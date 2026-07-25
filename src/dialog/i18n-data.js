// i18n-data.js - bilingual string tables for the add-in
//
// Loaded BEFORE i18n.js. Defines window.I18N_DATA = { zh: {...}, en: {...} }.
// Fallback chain: currentLang -> en -> key.
// Use {paramName} placeholders for dynamic values; pass {paramName: value} to i18n.t().

window.I18N_DATA = {
  zh: {
    // --- HTML head ---
    htmlTitle: 'R 角调整',

    // --- Header ---
    headerTitle: 'R 角调整',
    headerSubtitle: '精确设置圆角矩形的 R 角（厘米 / 百分比）',

    // --- Status card ---
    statusLabelSelection: '选区',
    statusLabelCurrentR: '当前 R 角',
    statusLabelLocked: '已锁定',
    statusReading: '读选中…',

    // --- Shape list ---
    selectedShapes: '选中的圆角矩形',
    rescan: '🔄 重新读',
    emptyShapes: '在 PPT 里框选形状后会出现在这里',
    nonRoundRect: '非圆角矩形',
    unnamed: '（未命名）',

    // --- R value input ---
    radiusValue: 'R 角数值',
    unitCm: '厘米',
    unitPercent: '百分比',
    hintRadiusRange: '范围 0 ~ 形状短边 / 2。超过一半显示为最大圆角。',
    placeholderRadius: '0.30',
    applyRadius: '应用 R 角',

    // --- Preset library ---
    presetLibrary: '预设库（5 槽位）',
    saveCurrentR: '+ 保存当前 R 角',
    presetAddTitle: '把当前选中的圆角矩形 R 角（或输入框值）存为预设',
    emptyPreset: '保存常用 R 角，一键应用。点击「+ 保存当前值」开始。',
    hintPresetSession: '预设纯内存保存，关掉 task pane 后失效（v1.2 升级为跨 session）。',

    // --- Pipette / style brush ---
    styleBrush: 'R 角样式刷',
    hintPipette: '点击吸取一个圆角矩形的 R 角，再点其他形状应用',
    pipettePickR: '吸取 R 角',
    pipetteSyncTitle: '勾选后样式刷会把源形状的「防误触」状态**覆盖**到目标（双向：源开启则目标开启；源未开启则目标也解除）',
    pipetteSyncLabel: '刷防误触状态',
    pipetteStateIdle: '空闲',

    // --- Layout mode ---
    layoutMode: '布局模式（v1.2）',
    emptyLayout: '在 PPT 里选 1+ 个圆角矩形（可以包含其他形状），然后在下面指定哪个是父、哪些是子。',
    emptyLayoutSetup: '在 PPT 里选 1+ 个圆角矩形',
    layoutLabelParent: '父',
    layoutLabelChildren: '子',
    layoutLabelRows: '行',
    layoutLabelCols: '列',
    layoutLabelRowsCols: '行 / 列',
    layoutLabelPadding: '边距',
    layoutLabelGutter: '间距',
    layoutLabelChildSize: '子尺寸',
    layoutLabelRLink: 'R 角联动',
    layoutAutoColsHint: '（自动 = 子数 ÷ 行）',
    layoutCoupledOneChild: '（只有 1 个子）',
    layoutCoupledFactorHint: '（自动 = {N} ÷ 行，N={N} 的因子: {factors}）',
    layoutHintActive: '已激活布局',
    layoutHintNotInSel: '（已不在选区）',
    layoutHintChildShape: '当前形状是布局子项',
    layoutHintStrict: '🔒 {count} 个启用了防误触，请先关闭后再建布局',
    layoutHintCanBuild: '✅ 可以建 {rows}×{cols} 布局',
    layoutChildrenCount: '{count} 个（{rows}×{cols}）',
    layoutChildInfo: '当前形状属于一个布局（父：{parent}）。修改父或脱离布局后可独立调整。',
    layoutRlinkSame: '等距（r = 父 R，45° 方向和边方向都等于 padding）',
    layoutRlinkSubtract: '层级感（r = 父 R − padding，子 R 角小一圈；45° 方向比 padding 窄约 40%）',
    layoutRlinkOff: '不联动（手动）',
    layoutWarnTooTight: '⚠️ 边距/间距太大，挤不下',
    layoutWarnNotEnough: '⚠️ 子形状不足（需要 {expected}，找到 {childCount}）',
    layoutPreviewWarn: '⚠️ 挤不下',
    layoutPreviewSize: '{w} × {h} cm',
    layoutPgLinkTitle: '锁链：开启后间距跟随边距变化（spacing = padding）',
    buildLayout: '建布局',
    detachLayout: '脱离布局',
    detachThisLayout: '脱离此布局',
    hintLayoutSetup: '选 1 父 + rows×cols 子后，「建布局」按钮启用。',
    layoutHintEmpty: '在 PPT 里选 1+ 个圆角矩形',
    layoutEmpty: '在 PPT 里选 1+ 个圆角矩形',
    layoutEmptySetup: '在 PPT 里选 1+ 个圆角矩形',
    layoutHintNotInSel: '（已不在选区）',
    layoutActive: '已激活布局',
    layoutChildrenCountFmt: '{count} 个（{rows}×{cols}）',
    layoutCoupledOneChild: '（只有 1 个子）',
    layoutCoupledFactorFmt: '（自动 = {N} ÷ 行，N={N} 的因子: {factors}）',
    layoutWarnTooTight: '⚠️ 边距/间距太大，挤不下',
    layoutWarnNotEnoughFmt: '⚠️ 子形状不足（需要 {expected}，找到 {childCount}）',
    layoutPreviewWarn: '⚠️ 挤不下',
    layoutPreviewSizeFmt: '{w} × {h} cm',
    layoutHintStrictFmt: '🔒 {count} 个启用了防误触，请先关闭后再建布局',
    layoutHintCanBuildFmt: '✅ 可以建 {rows}×{cols} 布局',
    layoutHintChildShape: '当前形状是布局子项',
    layoutChildInfoPre: '当前形状属于一个布局（父：',
    layoutChildInfoPost: '）。修改父或脱离布局后可独立调整。',

    // --- Lock / strict ---
    lockRadius: '使用数值固定 R 角',
    lockRadiusOff: '关闭使用数值固定 R 角',
    strictLock: '防误触',
    strictTitle: '开启后任何修改都不会改 R 角（task pane / PPT 内编辑都拒绝）',
    strictHint: '开启后任何修改都不改 R 角',
    hintLockCombined: '开启使用数值固定 R 角后 R 角按厘米值保持，PPT 内的任何编辑（拖控制柄等）都会被反算。开启防误触可让 task pane 内的按钮也拒绝改值。',
    reapplyLock: '重新应用锁定（针对当前选区）',

    // --- Footer ---
    footerVersion: 'v1.3 · 布局模式 + 样式刷精修 + 全 driver 化',

    // --- Debug log ---
    debugTitle: '🔧 调试日志（点击展开）',
    debugSmokeTitle: '跑遍所有 driver 方法，在 PPT 实际验证一遍',
    smokeTest: '🧪 Driver 烟囱测试',
    debugCopyTitle: '复制全部日志到剪贴板',
    copy: '📋 复制',
    debugClearTitle: '清空日志',
    clear: '🗑️ 清空',

    // --- Status messages (dynamic) ---
    shapeCount: '已选 {count} 个',
    radiusSingle: '{value} {unit}',
    radiusRange: '{min} ~ {max} {unit}',
    statusReadFailedFmt: '读失败：{error}',
    statusShapeCountFmt: '{count} 个',
    statusShapeCountMixedFmt: '{count} 个（混合）',
    statusCurrentRCmFmt: '{value} 厘米',
    statusCurrentRPercentFmt: '{value} %',
    statusCurrentRRangeCmFmt: '{min} ~ {max} 厘米',
    statusCurrentRRangePercentFmt: '{min} ~ {max} %',

    // --- Toasts: layout ---
    toastLayoutFailedFmt: '布局应用失败：{error}',
    toastLayoutAutoReducedFmt: '{warn} — 自动缩减到 {rows} 行',
    toastLayoutAppliedFmt: '✅ 布局已应用 {applied} 个子形状{failed}{rHint}{strictHint}',
    toastRowsRange: '行数范围 1~5',
    toastColsRange: '列数范围 1~5',
    toastSelectParent: '请在列表里指定一个父',
    toastChildrenInsufficientFmt: '子不足（需要 {need}，选了 {chosen}）',
    toastBuildLayoutFailedFmt: '建布局失败：{error}',
    toastNoChildrenWrittenFmt: '⚠️ 没写成功任何子形状：{warn}',
    toastLayoutBuiltFmt: '🎯 已建立 {rows}×{cols} 布局（{applied} 个子形状{failed}）',
    toastDetachFailedFmt: '脱离失败：{error}',
    toastDetached: '已脱离布局（子形状的位置/尺寸/R 角保留）',
    toastDetachedThis: '已脱离此布局',

    // --- Toasts: debug log ---
    toastCopiedFmt: '📋 已复制 {count} 行日志',
    toastCopiedFallbackFmt: '📋 已复制 {count} 行（fallback）',
    toastLogCleared: '🗑️ 调试日志已清空',
    toastSmokeTestDoneFmt: '🧪 烟囱测试完成：{pass} 通过 / {fail} 失败 — 看 debug 日志',
    toastSmokeTestFailedFmt: '🧪 烟囱测试失败：{error}',

    // --- Toasts: read / selection ---
    toastStaleLayoutsFmt: 'ℹ️ 检测到 {count} 个 layout 子已被删除/移走，下次应用布局时会自动清理',
    toastReadFailedFmt: '读选区失败: {error}',

    // --- Toasts: apply R / lock / strict ---
    toastSelectRoundRect: '请先在 PPT 里框选圆角矩形',
    toastInvalidR: '请输入有效的 R 角值',
    toastStrictBlocksFmt: '🔒 防误触已开启（{count} 个），不能改 R 角。先关掉防误触或解锁。',
    toastRUpdatedFmt: '✅ 已更新 {count} 个圆角矩形为 {value}{lockHint}',
    toastPartialSuccessFmt: '⚠️ {updated} 个成功，{failed} 个失败（可能不是圆角矩形）',
    toastApplyFailedFmt: '应用失败：{error}',
    toastNotRoundRect: '选中的形状都不是圆角矩形',
    toastStrictCannotEnableFmt: '无法开启防误触：{name} 当前 R 角未知或为 0',
    toastStrictEnabledFmt: '🔒 防误触已开启（{count} 个）— 已自动用当前 R 角作固定值',
    toastStrictDisabledFmt: '防误触已关闭（{count} 个）— 允许主动调整 R 角{keepHint}',
    toastNoLockedRects: '当前选区没有锁定的圆角矩形',
    toastReappliedFmt: '🔒 重新应用了 {count} 个锁定的 R 角{failed}',
    toastLockRecomputedFmt: '🔒 使用数值固定 R 角：反算了 {count} 个被改的 R 角',
    toastLockFollowedFmt: '🪄 使用数值固定 R 角已跟随 R 角滑块更新（{count} 个）',

    // --- Toasts: preset ---
    toastRenamedFmt: '已重命名为「{name}」',
    toastPercentRange: '百分比模式范围 0~50%',
    toastPresetUpdatedFmt: '已更新「{name}」为 {value}',
    toastPresetsFullFmt: '预设库已满（{max} 个），请先删一个再加',
    toastNoRValue: '请先在 PPT 里选 1 个圆角矩形，或在输入框里输入有效的 R 角值',
    toastPresetSavedFmt: '已保存为「{name}」({value})',
    toastPresetDeleted: '预设已删除',
    presetEditTitleFmt: '点击编辑数值（按 {unit} 解读），回车保存',

    // --- Toasts: pipette / brush ---
    toastNotRoundRectPickup: '当前选中的不是圆角矩形，请先在 PPT 里点 1 个圆角矩形',
    toastEnterPipette: '🎯 进入吸取模式 — 在 PPT 里点 1 个圆角矩形',
    toastPipetteClosed: '样式刷已关闭',
    toastPickupFailedFmt: '吸取失败：{error}',
    toastNoRoundRectInSelection: '选区里没有圆角矩形',
    toastPickedFmt: '🪣 已吸取「{name}」= {value} — 选中目标形状自动应用{strictHint}',
    pipetteHintSourcing: '在 PPT 里点选 1 个圆角矩形吸取其 R 角',
    toastBrushFailedFmt: '样式刷失败：{error}',
    toastBrushBlocked: '🔒 选区里有形状启用了防误触，样式刷不生效。先关掉目标的防误触或解锁。',
    toastBrushNoSelection: '选区为空，先在 PPT 里选 1 个圆角矩形',
    toastBrushAppliedFmt: '🪣 样式刷应用了 {count} 个圆角矩形{failed}{lockHint}{strictHint}',

    // --- Hints (suffix/optional) ---
    failedStrFmt: '，{count} 个失败',
    lockHint: '（已固定 R）',
    strictHint: '（已刷防误触状态）',
    keepLockHint: '，固定值保留',
    rLinkHint: '（含 R 角联动）',

    // --- ARIA ---
    ariaUnit: '单位',
    ariaHistory: '最近 5 次',
  },

  en: {
    // --- HTML head ---
    htmlTitle: 'RadiusInPpt',

    // --- Header ---
    headerTitle: 'RadiusInPpt',
    headerSubtitle: 'Set rounded rectangle corner radius precisely (cm / %)',

    // --- Status card ---
    statusLabelSelection: 'Selection',
    statusLabelCurrentR: 'Current R',
    statusLabelLocked: 'Locked',
    statusReading: 'Reading selection…',

    // --- Shape list ---
    selectedShapes: 'Selected rounded rectangles',
    rescan: '🔄 Re-read',
    emptyShapes: 'Select shapes in PowerPoint and they will appear here',
    nonRoundRect: 'Not a rounded rectangle',
    unnamed: '(unnamed)',

    // --- R value input ---
    radiusValue: 'R value',
    unitCm: 'cm',
    unitPercent: '%',
    hintRadiusRange: 'Range 0 ~ short side / 2. Values above half show as max corner.',
    placeholderRadius: '0.30',
    applyRadius: 'Apply R',

    // --- Preset library ---
    presetLibrary: 'Presets (5 slots)',
    saveCurrentR: '+ Save current R',
    presetAddTitle: 'Save the current selected rounded rectangle R (or input value) as a preset',
    emptyPreset: 'Save common R values for one-click apply. Click "+ Save current value" to start.',
    hintPresetSession: 'Presets are kept in memory only and are lost when the task pane is closed (cross-session planned for v1.2+).',

    // --- Pipette / style brush ---
    styleBrush: 'R style brush',
    hintPipette: 'Click to pick the R from a rounded rectangle, then click other shapes to apply',
    pipettePickR: 'Pick R',
    pipetteSyncTitle: 'When checked, the brush **overrides** the source shape\'s "anti-misclick" (strict) state on the targets (bidirectional: source on -> target on; source off -> target off).',
    pipetteSyncLabel: 'Sync strict-lock state',
    pipetteStateIdle: 'Idle',

    // --- Layout mode ---
    layoutMode: 'Layout mode (v1.2)',
    emptyLayout: 'Select 1+ rounded rectangles in PowerPoint (other shapes can be included), then designate which is the parent and which are children below.',
    emptyLayoutSetup: 'Select 1+ rounded rectangles in PowerPoint',
    layoutEmpty: 'Select 1+ rounded rectangles in PowerPoint (other shapes can be included), then designate which is the parent and which are children below.',
    layoutEmptySetup: 'Select 1+ rounded rectangles in PowerPoint',
    layoutHintEmpty: 'Select 1+ rounded rectangles in PowerPoint',
    layoutHintNotInSel: '(not in current selection)',
    layoutActive: 'Layout active',
    layoutChildrenCountFmt: '{count} ({rows}×{cols})',
    layoutCoupledOneChild: '(only 1 child)',
    layoutCoupledFactorFmt: '(auto = {N} ÷ rows, N={N} factors: {factors})',
    layoutWarnTooTight: '⚠️ Padding/gutter too large, no fit',
    layoutWarnNotEnoughFmt: '⚠️ Not enough children (need {expected}, found {childCount})',
    layoutPreviewWarn: '⚠️ No fit',
    layoutPreviewSizeFmt: '{w} × {h} cm',
    layoutHintStrictFmt: '🔒 {count} shapes have strict-lock enabled. Disable first before building a layout.',
    layoutHintCanBuildFmt: '✅ Can build {rows}×{cols} layout',
    layoutHintChildShape: 'Current shape is a layout child',
    layoutChildInfoPre: 'Current shape belongs to a layout (parent: ',
    layoutChildInfoPost: '). Modify the parent or detach to edit independently.',
    layoutLabelParent: 'Parent',
    layoutLabelChildren: 'Children',
    layoutLabelRows: 'Rows',
    layoutLabelCols: 'Cols',
    layoutLabelRowsCols: 'Rows / Cols',
    layoutLabelPadding: 'Padding',
    layoutLabelGutter: 'Gutter',
    layoutLabelChildSize: 'Child size',
    layoutLabelRLink: 'R coupling',
    layoutAutoColsHint: '(auto = children ÷ rows)',
    layoutCoupledOneChild: '(only 1 child)',
    layoutCoupledFactorHint: '(auto = {N} ÷ rows, N={N} factors: {factors})',
    layoutHintActive: 'Layout active',
    layoutHintNotInSel: '(not in current selection)',
    layoutHintChildShape: 'Current shape is a layout child',
    layoutHintStrict: '🔒 {count} shapes have strict-lock enabled. Disable first before building a layout.',
    layoutHintCanBuild: '✅ Can build {rows}×{cols} layout',
    layoutChildrenCount: '{count} ({rows}×{cols})',
    layoutChildInfo: 'Current shape belongs to a layout (parent: {parent}). Modify the parent or detach to edit independently.',
    layoutRlinkSame: 'Equal (r = parent R, 45° and side directions both equal padding)',
    layoutRlinkSubtract: 'Hierarchical (r = parent R − padding, child R one notch smaller; 45° direction ~40% narrower than padding)',
    layoutRlinkOff: 'No coupling (manual)',
    layoutWarnTooTight: '⚠️ Padding/gutter too large, no fit',
    layoutWarnNotEnough: '⚠️ Not enough children (need {expected}, found {childCount})',
    layoutPreviewWarn: '⚠️ No fit',
    layoutPreviewSize: '{w} × {h} cm',
    layoutPgLinkTitle: 'Chain link: when active, gutter follows padding (spacing = padding)',
    buildLayout: 'Build layout',
    detachLayout: 'Detach layout',
    detachThisLayout: 'Detach this layout',
    hintLayoutSetup: 'Select 1 parent + rows×cols children to enable "Build layout".',

    // --- Lock / strict ---
    lockRadius: 'Fix R by value',
    lockRadiusOff: 'Disable fix-R-by-value',
    strictLock: 'Anti-misclick (strict)',
    strictTitle: 'When enabled, all changes (task pane / in-PPT) are rejected and R does not change',
    strictHint: 'When enabled, all changes are rejected',
    hintLockCombined: 'When "Fix R by value" is on, R is held at the cm value and any in-PPT edits (handles, etc.) are reversed. Enable anti-misclick to also reject task-pane button changes.',
    reapplyLock: 'Re-apply lock (current selection)',

    // --- Footer ---
    footerVersion: 'v1.3 · Layout + style-brush polish + full driver-ification',

    // --- Debug log ---
    debugTitle: '🔧 Debug log (click to expand)',
    debugSmokeTitle: 'Run all driver methods, verify in real PowerPoint',
    smokeTest: '🧪 Driver smoke test',
    debugCopyTitle: 'Copy all logs to clipboard',
    copy: '📋 Copy',
    debugClearTitle: 'Clear log',
    clear: '🗑️ Clear',

    // --- Status messages (dynamic) ---
    shapeCount: '{count} selected',
    radiusSingle: '{value} {unit}',
    radiusRange: '{min} ~ {max} {unit}',
    statusReadFailedFmt: 'Read failed: {error}',
    statusShapeCountFmt: '{count}',
    statusShapeCountMixedFmt: '{count} (mixed)',
    statusCurrentRCmFmt: '{value} cm',
    statusCurrentRPercentFmt: '{value} %',
    statusCurrentRRangeCmFmt: '{min} ~ {max} cm',
    statusCurrentRRangePercentFmt: '{min} ~ {max} %',

    // --- Toasts: layout ---
    toastLayoutFailedFmt: 'Layout apply failed: {error}',
    toastLayoutAutoReducedFmt: '{warn} — auto-shrunk to {rows} rows',
    toastLayoutAppliedFmt: '✅ Layout applied to {applied} children{failed}{rHint}{strictHint}',
    toastRowsRange: 'Row count range 1~5',
    toastColsRange: 'Col count range 1~5',
    toastSelectParent: 'Please designate a parent in the list',
    toastChildrenInsufficientFmt: 'Not enough children (need {need}, selected {chosen})',
    toastBuildLayoutFailedFmt: 'Build layout failed: {error}',
    toastNoChildrenWrittenFmt: '⚠️ No children written: {warn}',
    toastLayoutBuiltFmt: '🎯 Built {rows}×{cols} layout ({applied} children{failed})',
    toastDetachFailedFmt: 'Detach failed: {error}',
    toastDetached: 'Detached from layout (child position / size / R preserved)',
    toastDetachedThis: 'Detached from this layout',

    // --- Toasts: debug log ---
    toastCopiedFmt: '📋 Copied {count} log lines',
    toastCopiedFallbackFmt: '📋 Copied {count} lines (fallback)',
    toastLogCleared: '🗑️ Debug log cleared',
    toastSmokeTestDoneFmt: '🧪 Smoke test done: {pass} pass / {fail} fail — see debug log',
    toastSmokeTestFailedFmt: '🧪 Smoke test failed: {error}',

    // --- Toasts: read / selection ---
    toastStaleLayoutsFmt: 'ℹ️ {count} layout child(ren) have been deleted / moved; will auto-clean on next layout apply',
    toastReadFailedFmt: 'Read selection failed: {error}',

    // --- Toasts: apply R / lock / strict ---
    toastSelectRoundRect: 'Please select rounded rectangles in PowerPoint first',
    toastInvalidR: 'Please enter a valid R value',
    toastStrictBlocksFmt: '🔒 Strict-lock enabled on {count} shape(s). R cannot be changed. Disable strict-lock or unlock first.',
    toastRUpdatedFmt: '✅ Updated {count} rounded rectangle(s) to {value}{lockHint}',
    toastPartialSuccessFmt: '⚠️ {updated} succeeded, {failed} failed (may not be rounded rectangles)',
    toastApplyFailedFmt: 'Apply failed: {error}',
    toastNotRoundRect: 'None of the selected shapes are rounded rectangles',
    toastStrictCannotEnableFmt: 'Cannot enable strict-lock: {name} has unknown or 0 R',
    toastStrictEnabledFmt: '🔒 Strict-lock enabled on {count} shape(s) — current R auto-used as fixed value',
    toastStrictDisabledFmt: 'Strict-lock disabled on {count} shape(s) — manual R edits allowed{keepHint}',
    toastNoLockedRects: 'No locked rounded rectangles in the current selection',
    toastReappliedFmt: '🔒 Re-applied {count} locked R{failed}',
    toastLockRecomputedFmt: '🔒 Fix-R-by-value: recomputed {count} edited R(s) back to fixed value',
    toastLockFollowedFmt: '🪄 Fix-R-by-value followed R slider edit ({count} updated)',

    // --- Toasts: preset ---
    toastRenamedFmt: 'Renamed to "{name}"',
    toastPercentRange: 'Percentage mode range: 0~50%',
    toastPresetUpdatedFmt: 'Updated "{name}" to {value}',
    toastPresetsFullFmt: 'Preset library full ({max}); delete one first',
    toastNoRValue: 'Please select a rounded rectangle in PowerPoint, or enter a valid R in the input',
    toastPresetSavedFmt: 'Saved as "{name}" ({value})',
    toastPresetDeleted: 'Preset deleted',
    presetEditTitleFmt: 'Click to edit value (interpreted as {unit}); Enter to save',

    // --- Toasts: pipette / brush ---
    toastNotRoundRectPickup: 'Current selection is not a rounded rectangle. Click a rounded rectangle in PowerPoint first.',
    toastEnterPipette: '🎯 Entered pickup mode — click a rounded rectangle in PowerPoint',
    toastPipetteClosed: 'Style brush closed',
    toastPickupFailedFmt: 'Pickup failed: {error}',
    toastNoRoundRectInSelection: 'No rounded rectangles in selection',
    toastPickedFmt: '🪣 Picked "{name}" = {value} — select target shapes to auto-apply{strictHint}',
    pipetteHintSourcing: 'Click a rounded rectangle in PowerPoint to pick its R',
    toastBrushFailedFmt: 'Brush failed: {error}',
    toastBrushBlocked: '🔒 Selection contains strict-locked shapes. Brush will not apply. Disable targets first.',
    toastBrushNoSelection: 'Selection empty. Select a rounded rectangle in PowerPoint first.',
    toastBrushAppliedFmt: '🪣 Brush applied to {count} rounded rectangle(s){failed}{lockHint}{strictHint}',

    // --- Hints (suffix/optional) ---
    failedStrFmt: ', {count} failed',
    lockHint: ' (R locked)',
    strictHint: ' (strict state synced)',
    keepLockHint: ', fixed value preserved',
    rLinkHint: ' (with R coupling)',

    // --- ARIA ---
    ariaUnit: 'Unit',
    ariaHistory: 'Last 5',
  }
};
