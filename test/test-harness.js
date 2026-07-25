/*
 * test-harness.js — driver harness：call tracker + snapshot + 自测工具
 *
 * 目的：
 *   业务方法（writeRadius / applyLayout / ...）通过 driver 间接操作 shape。
 *   这个 harness 把 driver 包装一层，让测试能：
 *     1. 记录所有 driver 方法调用（calls）
 *     2. 拿到当前所有 shape 的 snapshot（adjFraction / tags / box）
 *     3. 用 assertSnapshot 验证业务执行后状态对不对
 *
 * 为什么需要这个：
 *   - 之前 test 里每个 case 都自己造 makeMockShape + makeMockDriver，散落重复
 *   - 测试失败时排查困难（不知道 driver 被调了哪些方法、shape 变成啥样）
 *   - 之后加新功能，AI 写测试不知道从哪开始
 *
 * 用法：
 *   const { createHarness } = require('./test-harness');
 *   const h = createHarness({ shapes: [r1, r2, ...] });
 *
 *   // 调业务方法
 *   await writeRadius(h.driver, r1, 0.5);
 *
 *   // 验证 driver 反应
 *   h.assertCalled('setAdjFraction');         // setAdjFraction 至少被调 1 次
 *   h.assertNotCalled('addTag');             // addTag 没被调
 *   h.assertCallCount('sync', 1);            // sync 调了 1 次
 *
 *   // 验证 shape 状态
 *   h.assertShape(r1, { adjFraction: 0.236, tags: {} });
 *   h.snapshot();                            // 返回所有 shape 的当前状态
 *
 * 跟 createDriver 的关系：
 *   - harness 内部用 createDriver(ctx) 拿真实 driver API
 *   - 覆盖 addTag / deleteTag / readTag（用 mock 实现）
 *   - 覆盖 activeSlide / slideShapes（如果传了 opts.slide）
 *   - 其他方法（setAdjFraction / setBox / size / box / isRoundRect / adjFraction / loadAdjValue / load / sync）走真实 createDriver
 *   - 覆盖 load / sync 把调用记录到 calls（业务能看到 driver 的所有调用）
 */

const path = require('path');
const { createDriver } = require(path.join(__dirname, '..', 'src', 'lib', 'ppt-driver.js'));
const { makeStandardFixture, makeFixtureShape, PT_PER_CM, cm } = require('./fixtures');

/**
 * 创建 driver harness
 * @param {Object} [opts]
 *   - shapes: Array<shape>     所有 shape（喂给 slide）
 *   - activeSlideIndex: number 默认 0（多 slide 测试用）
 * @returns {Object} harness
 */
function createHarness(opts) {
  opts = opts || {};
  const shapes = opts.shapes || [];
  const activeSlideIndex = opts.activeSlideIndex || 0;

  // ── call tracker ──
  const calls = [];  // [{ method, args, time }]
  const recordCall = (method, args) => {
    calls.push({ method, args, time: Date.now() });
  };

  // ── ctx（mock）──
  const ctx = {
    sync: async () => { recordCall('sync', []); return; },
  };

  // ── slide（mock Office.js slide proxy）──
  const slide = {
    shapes: { items: shapes },
    load: (fields) => { recordCall('slide.load', [fields]); },
  };

  // ── driver ──
  // 直接用真实 createDriver 拿完整 API（确保 mock 跟真实 API 一致）
  const driver = createDriver(ctx);

  // 覆盖 load：记录调用（不调 realLoad，因为 mock proxy 没有真正的 load 行为）
  driver.load = (proxy, fields) => {
    recordCall('load', [{ proxy: proxy && proxy.id ? proxy.id : '(proxy)', fields }]);
    // mock proxy 不会真的记录 load 队列，反正我们的 mock shape 的所有字段都是同步可读的
  };

  // 覆盖 collection accessors
  driver.activeSlide = () => { recordCall('activeSlide', []); return slide; };
  driver.slideShapes = (s) => { recordCall('slideShapes', []); return s.shapes; };

  // 覆盖 addTag / deleteTag / readTag：mock tag 集合
  driver.addTag = (s, key, value) => {
    recordCall('addTag', [s.id, key, value]);
    s._tags[key] = String(value);
  };
  driver.deleteTag = (s, key) => {
    recordCall('deleteTag', [s.id, key]);
    delete s._tags[key];
  };
  // readTag 模拟真实 driver：会调 ctx.sync() 一次
  // （mock 里 sync 立刻 resolve，行为一致）
  driver.readTag = async (s, key) => {
    recordCall('readTag', [s.id, key]);
    await ctx.sync();
    return s._tags[key] != null ? s._tags[key] : null;
  };

  // 覆盖 setAdjFraction：调真实的 s.adjustments.set + 记录
  driver.setAdjFraction = (s, frac) => {
    recordCall('setAdjFraction', [s.id, frac]);
    s.adjustments.set(0, frac);
  };

  // 覆盖 setBox：直接写 4 个字段 + 记录
  driver.setBox = (s, box) => {
    recordCall('setBox', [s.id, box]);
    s.left = box.left;
    s.top = box.top;
    s.width = box.width;
    s.height = box.height;
  };

  // 覆盖 loadAdjValue：mock 不需要真的 load（adjustments.get(0) 同步），只记录
  driver.loadAdjValue = (s) => {
    recordCall('loadAdjValue', [s.id]);
    // mock: 不需要 load，下一次 adjustments.get(0).value 直接拿 _adjFraction
  };

  // 覆盖 shapeId：记录读 id（让测试能验证"没读 id"）
  driver.shapeId = (s) => {
    recordCall('shapeId', [s.id]);
    return s.id;
  };

  // 覆盖 size：记录读 size
  driver.size = (s) => {
    recordCall('size', [s.id]);
    return { width: s.width, height: s.height };
  };

  // 覆盖 box：记录读 box
  driver.box = (s) => {
    recordCall('box', [s.id]);
    return { left: s.left, top: s.top, width: s.width, height: s.height };
  };

  // 覆盖 isRoundRect：记录读
  driver.isRoundRect = (s) => {
    recordCall('isRoundRect', [s.id]);
    return s.adjustments.count > 0;
  };

  // 覆盖 adjFraction：记录读
  driver.adjFraction = (s) => {
    recordCall('adjFraction', [s.id]);
    if (s.adjustments.count === 0) return 0;
    try {
      return s.adjustments.get(0).value;
    } catch (_) {
      return 0;
    }
  };

  // ── snapshot 工具 ──
  function snapshot() {
    const result = {};
    for (const s of shapes) {
      result[s.id] = {
        id: s.id,
        width: s.width,
        height: s.height,
        left: s.left,
        top: s.top,
        adjFraction: s._adjFraction,
        tags: Object.assign({}, s._tags),
      };
    }
    return result;
  }

  // ── assertion helpers ──
  function assertCalled(method, opts) {
    opts = opts || {};
    const matches = calls.filter((c) => c.method === method);
    if (opts.with) {
      // 验证至少一个 call 包含 with 的参数
      const found = matches.some((c) => {
        return opts.with.every((arg, i) => {
          if (typeof arg === 'function') return arg(c.args[i]);
          return c.args[i] === arg;
        });
      });
      if (!found) {
        throw new Error(
          `assertCalled: 期望 driver.${method} 被以 ${JSON.stringify(opts.with)} 调一次，但没找到\n` +
          `实际 calls: ${JSON.stringify(calls, null, 2)}`
        );
      }
      return;
    }
    if (matches.length === 0) {
      throw new Error(
        `assertCalled: 期望 driver.${method} 被调至少 1 次，但没找到\n` +
        `实际 calls: ${JSON.stringify(calls, null, 2)}`
      );
    }
  }

  function assertNotCalled(method) {
    const matches = calls.filter((c) => c.method === method);
    if (matches.length > 0) {
      throw new Error(
        `assertNotCalled: 期望 driver.${method} 没被调，但调了 ${matches.length} 次\n` +
        `calls: ${JSON.stringify(matches, null, 2)}`
      );
    }
  }

  function assertCallCount(method, expectedCount) {
    const matches = calls.filter((c) => c.method === method);
    if (matches.length !== expectedCount) {
      throw new Error(
        `assertCallCount: 期望 driver.${method} 被调 ${expectedCount} 次，实际 ${matches.length} 次\n` +
        `calls: ${JSON.stringify(matches, null, 2)}`
      );
    }
  }

  function assertShape(shape, expected) {
    const actual = {
      adjFraction: shape._adjFraction,
      tags: Object.assign({}, shape._tags),
    };
    if (expected.adjFraction != null) {
      if (typeof expected.adjFraction === 'number') {
        if (Math.abs(actual.adjFraction - expected.adjFraction) > 1e-6) {
          throw new Error(
            `assertShape ${shape.id}: 期望 adjFraction=${expected.adjFraction}，实际 ${actual.adjFraction}`
          );
        }
      } else if (expected.adjFraction instanceof Function) {
        if (!expected.adjFraction(actual.adjFraction)) {
          throw new Error(
            `assertShape ${shape.id}: adjFraction=${actual.adjFraction} 不满足谓词`
          );
        }
      }
    }
    if (expected.tags) {
      for (const k of Object.keys(expected.tags)) {
        const ev = expected.tags[k];
        if (ev === undefined) {
          if (k in actual.tags) {
            throw new Error(`assertShape ${shape.id}: 期望 tags.${k} 不存在，实际 '${actual.tags[k]}'`);
          }
        } else if (ev instanceof Function) {
          if (!ev(actual.tags[k])) {
            throw new Error(`assertShape ${shape.id}: tags.${k}='${actual.tags[k]}' 不满足谓词`);
          }
        } else {
          if (actual.tags[k] !== ev) {
            throw new Error(
              `assertShape ${shape.id}: 期望 tags.${k}='${ev}'，实际 '${actual.tags[k]}'`
            );
          }
        }
      }
    }
    if (expected.box) {
      for (const field of ['left', 'top', 'width', 'height']) {
        if (expected.box[field] != null) {
          const ev = expected.box[field];
          const av = shape[field];
          if (typeof ev === 'number' && Math.abs(av - ev) > 1e-6) {
            throw new Error(
              `assertShape ${shape.id}: 期望 ${field}=${ev}，实际 ${av}`
            );
          } else if (typeof ev === 'function' && !ev(av)) {
            throw new Error(
              `assertShape ${shape.id}: ${field}=${av} 不满足谓词`
            );
          }
        }
      }
    }
  }

  function dumpCalls(filter) {
    const list = filter ? calls.filter((c) => c.method === filter) : calls;
    console.log(`[harness] calls (${list.length}):`);
    for (const c of list) {
      console.log(`  ${c.method}(${c.args.map((a) => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(', ')})`);
    }
  }

  function reset() {
    calls.length = 0;
  }

  return {
    driver,
    calls,
    shapes,
    ctx,
    slide,
    snapshot,
    assertCalled,
    assertNotCalled,
    assertCallCount,
    assertShape,
    dumpCalls,
    reset,
  };
}

/**
 * 简易测试框架（test() + after() 钩子）
 * 用法：
 *   const t = createTestRunner();
 *   t.test('name 1', () => { ... });
 *   t.test('name 2', async () => { ... });
 *   await t.run();
 */
function createTestRunner() {
  const tests = [];
  const setupFns = [];
  const teardownFns = [];

  function test(name, fn) {
    tests.push({ name, fn });
  }

  function beforeEach(fn) {
    setupFns.push(fn);
  }

  function afterEach(fn) {
    teardownFns.push(fn);
  }

  async function run() {
    let passed = 0, failed = 0;
    for (const { name, fn } of tests) {
      for (const s of setupFns) await s();
      try {
        await fn();
        passed++;
        console.log(`  ✓ ${name}`);
      } catch (e) {
        failed++;
        console.log(`  ✗ ${name}`);
        console.log(`     ${e.message}`);
        if (e.stack) console.log(`     ${e.stack.split('\n').slice(1, 4).join('\n     ')}`);
      }
      for (const t of teardownFns) await t();
    }
    console.log('\n' + '='.repeat(50));
    console.log(`结果: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(50));
    if (failed > 0) process.exit(1);
  }

  return { test, beforeEach, afterEach, run };
}

module.exports = {
  createHarness,
  createTestRunner,
  makeStandardFixture,
  makeFixtureShape,
  PT_PER_CM,
  cm,
};
