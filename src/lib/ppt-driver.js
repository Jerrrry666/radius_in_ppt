/*
 * ppt-driver.js — 与 PowerPoint Office.js 的唯一交互层
 *
 * 这是整个系统的**交互层**（见 AGENTS.md 2.2）：
 *   - 所有 Office.js 调用都在这里
 *   - 不知道任何业务概念（strict / lock / layout / pipette）
 *   - 每个方法假定 caller 已经 load + sync 过对应字段
 *   - PPT 验过一次没 bug 后不再修改
 *
 * 使用方法：
 *   await PowerPoint.run(async (ctx) => {
 *     const driver = createDriver(ctx);
 *     const sel = driver.selectedShapes();
 *     driver.load(sel, 'items/id, items/width, items/height, items/adjustments');
 *     await driver.sync();
 *     for (const s of sel.items) {
 *       driver.setAdjFraction(s, 0.05);
 *     }
 *     await driver.sync();
 *   });
 *
 * Mock 测试：
 *   const driver = createMockDriver();  // 假 ctx + 假 shapes
 *   直接喂给 radius-core.writeRadius(driver, shape, cm, opts)
 */

function createDriver(ctx) {
  return {
    // ── 加载 + 同步 ─────────────────────────────────────
    // 把 fields（'items/id, items/adjustments'）加到 proxy 的加载队列
    // 必须在 await sync() 之后读 proxy 字段
    load(proxy, fields) {
      proxy.load(fields);
    },
    // 触发 sync，必需 await
    sync() {
      return ctx.sync();
    },

    // ── Collection accessors（返回 Office.js proxy）────
    // 当前选中的形状
    selectedShapes() {
      return ctx.presentation.getSelectedShapes();
    },
    // 当前激活的 slide
    activeSlide() {
      return ctx.presentation.getSelectedSlides().getItemAt(0);
    },
    // 给定 slide 上的所有 shapes 集合
    slideShapes(slide) {
      return slide.shapes;
    },

    // ── 读（假定已 load + sync）───────────────────────
    shapeId(s) {
      return s.id;
    },
    // 形状的尺寸（width, height）—— caller 需要 load 过 width + height
    // 适合只关心大小、不关心位置的场景（如 R 角 clamp、检测 minSide）
    size(s) {
      return { width: s.width, height: s.height };
    },
    // 形状的完整 box（left, top, width, height）—— caller 需要 load 过 4 个字段
    // 适合需要定位的场景（如 layout 算子位置/尺寸）
    // 用这个方法时 caller 一定要在 load 里加 'items/left, items/top'
    box(s) {
      return { left: s.left, top: s.top, width: s.width, height: s.height };
    },
    isRoundRect(s) {
      return s.adjustments.count > 0;
    },
    // 返回 0~1 分数（Mac LTSC），不是 0~50000
    // 注意：caller 必须显式 load 'items/adjustments/items/value' + sync（v1.2.5 实测坑），
    // 否则 s.adjustments.get(0).value 会抛"尚未加载结果对象的值"——这里 try/catch 兜底返回 0
    adjFraction(s) {
      if (s.adjustments.count === 0) return 0;
      try {
        return s.adjustments.get(0).value;
      } catch (_) {
        // value 没 load，老老实实返回 0 而不是 throw（driver API 不 throw 契约）
        return 0;
      }
    },

    // ── 写（假定已 load）──────────────────────────────
    // per-shape 显式 load adjustments value（Mac LTSC task pane 必加，v1.2.5 实测坑）
    // collection-level load 'items/adjustments' 只填 .count，不填 .value
    // 用法：先 collection-level load + sync，再 per-shape 调这个 + 再 sync
    loadAdjValue(s) {
      s.adjustments.load('items/value');
    },

    setBox(s, box) {
      s.left = box.left;
      s.top = box.top;
      s.width = box.width;
      s.height = box.height;
    },
    setAdjFraction(s, frac) {
      s.adjustments.set(0, frac);
    },

    // ── Tag 操作（add/delete 不需要 load）────────────
    // 字符串化 value（OOXML 要求 string）
    addTag(s, key, value) {
      s.tags.add(key, String(value));
    },
    deleteTag(s, key) {
      s.tags.delete(key);
    },
    // 读 tag：async，因为要 load('value') + sync
    // 不存在时返回 null（不会 throw）
    readTag: async (s, key) => {
      try {
        const t = s.tags.getItem(key);
        t.load('value');
        await ctx.sync();
        return t.value;
      } catch (_) {
        return null;
      }
    },
    // 批量读所有 shape 的 tags dict（一次性返回，不在内部 sync）
    // **避免 readTag 的 per-call sync 在 for 循环内累积**（Mac LTSC 真实跑过会丢后几个 shape）
    // 用法：先 driver.load(slide, 'shapes/items/tags') + ctx.sync() → driver.readTagsBulk(items) → 一次拿全部
    readTagsBulk: (shapesArr) => {
      const result = {};
      if (!shapesArr) return result;
      const list = Array.isArray(shapesArr) ? shapesArr : (shapesArr.items || []);
      for (const s of list) {
        if (!s || !s.tags) continue;
        // 真实 PPT 上：s.tags 已经是 collection-level load 过的（caller 之前 sync 过）
        // 这里读所有 tag 的 value（需要每个 tag .value 已填，依赖之前的 collection load + sync）
        // 简化：直接收集所有 tag
        result[s.id] = {};
        try {
          // s.tags.items 存在时遍历（Office.js collection 有 .items）
          if (s.tags.items) {
            for (const t of s.tags.items) {
              if (t && t.key != null) {
                result[s.id][t.key] = t.value;
              }
            }
          }
        } catch (_) {}
      }
      return result;
    },
  };
}

// 暴露给 Node.js 测试 / browser 全局
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { createDriver };
}
if (typeof window !== 'undefined') {
  window.PptDriver = { createDriver };
}
