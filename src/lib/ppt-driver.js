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
  // ── GroupShape 处理（v1.3.1）───────────────────────────
  // Office.js 1.8+ 暴露 PowerPoint.ShapeGroup（getSelectedShapes 选组合时返回 group proxy，
  // 不会自动展平），业务层需要把 group 递归展平成叶子 shape 才能正常读写 R 角。
  //
  // 防误触铁律：group 本身不是 roundRect（adjustments.count=0），如果 caller 不展平
  // 会被现有的 isRoundRect 判定自动过滤掉，导致整个 group 在 UI 上"看不见"。
  //
  // 防死循环：group 理论上可嵌套（group 里再 group），用 seen Set 记已访问 shape id。
  //
  // Mock 兼容：测试 driver 对象可挂 _isGroup / _groupShapes 字段模拟 group 行为。

  // 判定 shape 是不是 group
  // 真实 PPT：s.type === 'Group'（= PowerPoint.ShapeType.group）
  // Mock：caller 挂 s._isGroup = true
  //
  // v1.3.1 Mac LTSC 实测（user log）：
  //   - 选 group → s.type = 'Group' 字符串（不是数字 enum）
  //   - 选普通 shape → s.type = 'GeometricShape'
  //   - 之前加的 `s.group && s.group.shapes` 兜底**反而误判**普通 roundRect
  //     为 group：因为 sel.load('items/.../items/group/shapes/items/...') 嵌套
  //     load 路径让 Office.js 在**非 group 节点**上也填出 s.group proxy（空，
  //     shapes 存在但 items 为空），触发兜底命中 → 普通 shape 被当 group
  //   - 修法：去掉兜底，只信 s.type
  const isGroup = (s) => {
    if (!s) return false;
    if (s._isGroup !== undefined) return !!s._isGroup;
    try {
      const t = s.type;
      // 兼容：'Group' 字符串 / PowerPoint.ShapeType.group 枚举 / msoGroup=6 数字
      if (t === 'Group' || t === 'GroupShape' || t === PowerPoint.ShapeType.group) return true;
      if (t === 6) return true;
      return false;
    } catch (_) {
      return false;
    }
  };

  // 取 group 的子 shape 数组（同步，不调 sync）
  // 真实 PPT：s.group.shapes.items（PowerPoint.ShapeGroup.shapes 是 ShapeScopedCollection）
  // Mock：caller 挂 s._groupShapes = [shape1, shape2, ...]
  const groupShapes = (g) => {
    if (!g) return [];
    if (g._groupShapes !== undefined) return Array.isArray(g._groupShapes) ? g._groupShapes : [];
    try {
      const scoped = g.group ? g.group.shapes : null;
      if (!scoped) return [];
      return scoped.items || [];
    } catch (_) {
      return [];
    }
  };

  // 递归展平 selectedShapes → 叶子 shape 数组（同步版）
  // - depth-first 遍历
  // - 跳过 group 节点本身，只把 group 内的子 shape 加到结果
  // - 嵌套 group 全部展平（group 里再 group）
  // - 用 seen Set 记已访问 shape id，防恶意循环引用
  // - 输入：getSelectedShapes() 返回的 ShapeScopedCollection 或普通数组
  // - 输出：普通数组（不是 Office.js collection，没有 .items 字段）
  //
  // v1.3.1 Mac LTSC 实测：
  //   - group 子集合必须先 load + sync，才能同步展平
  //   - 不要把 items/group/shapes/... 路径无条件 load 到普通 shape；单选普通
  //     GeometricShape 时 Mac LTSC 会在 sync 抛 GeneralException
  //   - loadShapeTree 负责先读 type，再只对真实 Group 的 shapes collection 做 collection-level load
  // 最近一次 flatten/loadShapeTree 建出的「子 id → 直接父 group」索引。
  // 不使用 Shape.parentGroup：Mac LTSC 对顶层 shape 访问 parentGroup 会抛
  // GeneralException；遍历已确认的 group.shapes 更安全。
  const parentGroupById = new Map();

  const flattenSelected = (selectedShapes) => {
    const list = Array.isArray(selectedShapes)
      ? selectedShapes
      : (selectedShapes && selectedShapes.items) || [];
    const out = [];
    const seen = new Set();
    parentGroupById.clear();
    const walk = (s, parentGroup) => {
      if (!s) return;
      let id = null;
      try { id = s.id; } catch (_) {}
      if (id != null && parentGroup) parentGroupById.set(id, parentGroup);
      if (id != null) {
        if (seen.has(id)) return;  // 防循环
        seen.add(id);
      }
      if (isGroup(s)) {
        for (const sub of groupShapes(s)) walk(sub, s);
      } else {
        out.push(s);
      }
    };
    for (const s of list) walk(s, null);
    return out;
  };

  // 选区顶层是否包含 group（不递归）。
  // 用于区分「用户拖整个 group」和「用户进入 group 后拖单个叶子」：
  // 前者 PowerPoint 已经原生缩放全部后代，caller 不应再逐个重写子 box。
  const hasTopLevelGroup = (selectedShapes) => {
    const list = Array.isArray(selectedShapes)
      ? selectedShapes
      : (selectedShapes && selectedShapes.items) || [];
    return list.some(isGroup);
  };

  // shape 所在的 group 层级（PowerPointApi 1.8+）。
  // 0 = 顶层普通 shape / 顶层 group；1+ = group 内部 shape。
  // caller 必须先 load level。Mock 可用 _groupLevel。
  const shapeLevel = (s) => {
    if (!s) return 0;
    if (s._groupLevel !== undefined) {
      const mockLevel = Number(s._groupLevel);
      return Number.isFinite(mockLevel) && mockLevel > 0 ? mockLevel : 0;
    }
    try {
      const level = Number(s.level);
      return Number.isFinite(level) && level > 0 ? level : 0;
    } catch (_) {
      return 0;
    }
  };

  // 取最近一次 shape tree 遍历记录的直接父 group；顶层 shape 返回 null。
  const parentGroupOf = (s) => {
    if (!s) return null;
    let id = null;
    try { id = s.id; } catch (_) {}
    return id != null ? (parentGroupById.get(id) || null) : null;
  };

  // 取最外层 group。只沿安全的 tree 索引走，不访问 Shape.parentGroup。
  const topGroupOf = (s) => {
    let parent = parentGroupOf(s);
    if (!parent) return null;
    const seen = new Set();
    let top = parent;
    while (parent) {
      let id = null;
      try { id = parent.id; } catch (_) {}
      const key = id != null ? `id:${id}` : parent;
      if (seen.has(key)) break;
      seen.add(key);
      top = parent;
      parent = parentGroupOf(parent);
    }
    return top;
  };

  // 加载一个 shape collection 的完整树并返回叶子数组。
  //
  // 为什么不用 collection.load('items/group/shapes/items/...')：
  // Mac LTSC 在选区含普通 shape 时会把 group path 应用到非 group 节点，ctx.sync()
  // 直接抛 GeneralException。这里分阶段处理：
  //   1. 顶层 collection 只 load 普通字段 + type
  //   2. sync 后按 type 找真实 Group
  //   3. 只对 group.shapes collection 做下一层 collection-level load
  //   4. 重复直到没有嵌套 group
  //
  // fields 使用 shape 字段名（如 'id, width, height, adjustments, tags'），
  // 方法会自动补 id + type。返回普通叶子数组。
  const loadShapeTree = async (shapeCollection, fields) => {
    if (!shapeCollection) return [];
    if (Array.isArray(shapeCollection) || typeof shapeCollection.load !== 'function') {
      return flattenSelected(shapeCollection);
    }

    const requested = String(fields || '')
      .split(/,\s*/)
      .map((f) => f.trim())
      .filter(Boolean);
    const allFields = Array.from(new Set(['id', 'type', ...requested]));
    const itemFields = allFields.map((f) => `items/${f}`).join(', ');
    const loadCollection = (collection) => collection.load(itemFields);

    loadCollection(shapeCollection);
    await ctx.sync();

    let frontier = (shapeCollection.items || []).filter(isGroup);
    const expanded = new Set();
    while (frontier.length > 0) {
      const loadedGroups = [];
      for (const groupShape of frontier) {
        let id = null;
        try { id = groupShape.id; } catch (_) {}
        const key = id != null ? `id:${id}` : groupShape;
        if (expanded.has(key)) continue;
        expanded.add(key);
        try {
          const children = groupShape.group && groupShape.group.shapes;
          if (!children || typeof children.load !== 'function') continue;
          loadCollection(children);
          loadedGroups.push(children);
        } catch (_) {
          // type 已确认是 Group，但 host 没给出 group.shapes 时安全跳过
        }
      }
      if (loadedGroups.length === 0) break;
      await ctx.sync();
      frontier = [];
      for (const children of loadedGroups) {
        for (const child of children.items || []) {
          if (isGroup(child)) frontier.push(child);
        }
      }
    }

    return flattenSelected(shapeCollection);
  };

  // 读取已经 load + sync 完成的 TagCollection。
  // PowerPoint 会把 tag key 统一存成大写；driver 保留宿主返回的原始 key，
  // 业务层需要按大小写不敏感方式查找。
  const readTagsBulk = (shapesArr) => {
    const result = {};
    if (!shapesArr) return result;
    const list = Array.isArray(shapesArr) ? shapesArr : (shapesArr.items || []);
    for (const s of list) {
      if (!s || !s.tags) continue;
      let id = null;
      try { id = s.id; } catch (_) {}
      if (id == null) continue;
      result[id] = {};
      try {
        if (s.tags.items) {
          for (const t of s.tags.items) {
            if (t && t.key != null) {
              result[id][t.key] = t.value;
            }
          }
        }
      } catch (_) {}
    }
    return result;
  };

  // 批量加载所有目标 shape 的 TagCollection key/value，一次 ctx.sync 后返回 dict。
  //
  // 不能只在 shape collection 上 load('items/tags')：那只加载导航属性，
  // 不会填充 s.tags.items。也不能在业务循环里逐 shape readTag + sync：
  // Mac LTSC 会让后续 adjustment 写入丢失。TagCollection 官方契约是
  // collection.load('key, value')。
  const loadTagsBulk = async (shapesArr) => {
    const list = Array.isArray(shapesArr)
      ? shapesArr
      : (shapesArr && shapesArr.items) || [];
    let queued = false;
    for (const s of list) {
      if (!s || !s.tags || typeof s.tags.load !== 'function') continue;
      s.tags.load('key, value');
      queued = true;
    }
    if (queued) await ctx.sync();
    return readTagsBulk(list);
  };

  return {
    // ── 加载 + 同步 ─────────────────────────────────────
    // 把 fields（'items/id, items/adjustments'）加到 proxy 的加载队列
    // 必须在 await sync() 之后读 proxy 字段
    load(proxy, fields) {
      proxy.load(fields);
    },
    loadShapeTree,
    loadTagsBulk,
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

    // ── GroupShape（v1.3.1）─────────────────────────────
    isGroup,
    groupShapes,
    flattenSelected,
    hasTopLevelGroup,
    shapeLevel,
    parentGroupOf,
    topGroupOf,

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
    // 解除一个已确认的 group（PowerPointApi 1.8+）
    ungroupShapeGroup(s) {
      s.group.ungroup();
    },
    // 把 shape id / proxy 数组重新组合（PowerPointApi 1.8+）
    addGroup(shapeCollection, values) {
      return shapeCollection.addGroup(values);
    },
    selectShapes(slide, shapeIds) {
      slide.setSelectedShapes(shapeIds);
    },
    shapeName(s) {
      return s.name;
    },
    setShapeName(s, name) {
      s.name = name;
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
    // 同步读取已加载的 tags；多数业务应优先用 loadTagsBulk。
    readTagsBulk,
  };
}

// 暴露给 Node.js 测试 / browser 全局
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { createDriver };
}
if (typeof window !== 'undefined') {
  window.PptDriver = { createDriver };
}
