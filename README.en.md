# RadiusInPpt — macOS PowerPoint Add-in

> 🌍 **[中文版](./README.md)** · This is the English version. The Chinese README is the canonical one.

A custom **"R 角调整"** tab in PowerPoint's ribbon (a la iSlide) that lets you set the corner radius of rounded rectangles precisely in **centimeters** or **percentages**, with absolute-value locking, anti-misclick, a style brush, 5-step history, and the **v1.2 Layout Mode** (rows × cols grid + padding/gutter sliders + R-corner coupling).

## 📌 v1.3 Update (2026-07-26)

v1.2 finished "nested rounded-rectangle distribution." v1.3 polishes the layout / style-brush details, completes the dialog.js / radius-core full driver migration, and finishes Step 3-4 migration.

**Layout mode refinements**:
- **Row/column coupled slider** — one slider is enough; columns = children ÷ rows (rows × cols = N strictly, no empty slots)
- **Row values are now a discrete list** — positive factors of N (datalist tick hints), e.g. N=4 → [1, 2, 4], no "3×2=6 → 5 empty slots" cases
- **Padding / gutter Photoshop-style chain link** — chain icon centered between two rows; when active, gutter = padding
- **Gutter disabled when chained** — full greyed-out + non-interactive (eliminates the "gutter modified → chain reverts → shape not reverted" race)

**Style brush strict bidirectional override**:
- **"Apply strict-lock state" = bidirectional override** — source strict=true → target strict=true; source strict=false → target strict=false
- Order matters: source=true → **write R first, then add strict** (avoids writeRadius being blocked); source=false → **delete strict first, then write R** (lets writeRadius through)

**Architecture cleanup** (v1.3.0):
- **dialog.js / radius-core fully driver-ified** — 8 driver-version functions replace scattered ctxShape operations
- **Step 3-4 migration complete** — layout tag read/write + pipette go through driver
- **3 long-standing bugs fixed** — style-brush pick-up not applying / layout R-coupling writing only 2 of 4 children / lockMonitor `GeneralException`
- **AGENTS.md §1.0 new rule** — AI commit/push permission limit (lesson: v1.2.14 first-deploy wrong position → force-push to rewrite public history)
- **AGENTS.md §1.0 新规则** — AI commit/push 权限限制

**Tests 210/0** — 95 features + 115 radius-core (6 new tests added for bidirectional strict override)

Full changelog: [`changelogs/v1.3.md`](./changelogs/v1.3.md) (covers v1.2.8 → v1.3.0) · History: [`changelogs/v1.2.md`](./changelogs/v1.2.md) · Main log: [`LOG.md`](./LOG.md) · Roadmap: [`plans/feature-roadmap.md`](./plans/feature-roadmap.md)

## The Problem It Solves

PowerPoint's built-in "Rounded Rectangle" shape has these pain points:
- Corner radius is a **relative value** (0% ~ 50% of the short side), not absolute
- When you resize the shape, the R-corner scales with it
- The "Format Shape" panel has no direct cm input
- **Stacking multiple rounded rectangles** means manually computing position / size / R for every child

This add-in lets you:
- ✅ Enter `0.3 cm` as an **absolute value** for the R-corner
- ✅ Toggle **lock** — when locked, R stays in cm and re-scales proportionally when the shape resizes
- ✅ Act on **multiple selected rounded rectangles** at once
- ✅ v1.2: 1 parent + N children → drag a slider → real-time even distribution + R-coupling

## User Experience

Distributed as a macOS `.app` (~440 KB), double-click to launch:

1. Double-click `R 角调整.app` (or `RadiusInPpt.app`)
2. A prompt appears → choose "Quit and reopen PowerPoint"
3. Reopen PowerPoint → **"R 角调整"** tab appears in the ribbon
4. Click **"调整 R 角"** in the tab → task pane opens on the right
5. Select a rounded rectangle → type `0.3` (cm) or `10` (%) in the task pane → Apply / Lock / Strict
6. **v1.2**: select 1 parent + N children → Group → "Build Layout" → drag sliders to distribute in real time

> After that, just double-click the `.app` every time you need it (server runs in background, manifest is persistent).
> Note: if you changed code, fully quit PowerPoint (`Cmd+Q`) and reopen for the task pane to pick up new code.

## Features

| Feature | Description |
| --- | --- |
| Custom ribbon tab | "R 角调整" tab + "调整 R 角" button (iSlide-style position) |
| Task pane | 360×560 sidebar, modeless, doesn't block view |
| cm / percentage input | `cm` ↔ `%` toggle; % is interpreted as fraction of short side; enter + "Apply R" or Enter key |
| **v1.2** Layout mode | 1 parent + N children, rows×cols grid + padding/gutter sliders + R coupling |
| **v1.2** R coupling | child R = `max(0, parentR − padding)`; off / same / subtract modes |
| **v1.2** Nested state persistence | parent stores JSON + children store parentShapeId as bidirectional tags; travels with the .pptx |
| **v1.3** Row/column coupled slider | one slider, columns = children ÷ rows (rows × cols = N strictly) |
| **v1.3** Row discrete factor list | valid row values = positive factors of N ([1, 2, 4] / [1, 2, 3, 6] / prime [1, N]), no empty slots |
| **v1.3** Padding/gutter chain link | Photoshop-style chain icon; when active, gutter = padding; gutter fully disabled when chained (avoids race conditions) |
| **v1.3** Style brush strict bidirectional | checkbox toggles bidirectional override: source strict=true → target strict=true; source strict=false → target strict=false |
| Fixed R by value | button on/off; when on, resizing in PPT re-computes to fixed value |
| Anti-misclick (strict lock) | independent toggle; when on, uses current R as fixed value; rejects task pane edits + reverses R-slider drags |
| R preset library | 5 user-editable presets, name + value, one-click apply |
| R style brush | pick from 1 roundRect's R, paint to other targets; optional "apply strict-lock state" bidirectional override |
| Multi-select | acts on all selected rounded rectangles; non-roundRects are skipped with a notice |
| Shape list | live display of each selected shape's current R (syncs with in-PPT edits) |
| Auto-reapply on lock | setInterval polling, distinguishes "resize drag" from "R-slider drag" |
| Persistence | lock info stored on each shape's `shape.tags` (OOXML `<p:tagLst>`), survives across devices |

## Project Structure

```
radius_in_ppt/
├── manifest.xml                       # Office Add-in manifest (points to localhost:3000)
├── src/
│   ├── dialog/                        # task pane UI ("dialog" is a historical name)
│   │   ├── dialog.html
│   │   ├── dialog.js                  # ~2580 lines (post-v1.3 cleanup; Step 5 refactor target ~500)
│   │   └── dialog.css
│   └── lib/                           # implementation + interaction layers (v1.2)
│       ├── radius-core.js             # ~1470 lines, pure algorithms + 10 driver-version feature functions
│       └── ppt-driver.js              # 109 lines, 16 Office.js interaction methods
├── app/MacOS/RadiusInPpt              # bash launcher
├── tools/
│   ├── serve.js                       # ~60 lines static file server
│   ├── build-app.sh                   # package as .app
│   ├── build-and-deploy.sh            # one-click build + deploy + git commit
│   ├── build-dmg.sh                   # optional: package as .dmg
│   └── sign-and-notarize.sh           # optional: code sign + notarize
├── assets/                            # ribbon icons (5 sizes, referenced by manifest.xml)
├── test/                              # unit tests (210 total)
│   ├── test-radius-core.js            #   115 pure algorithm
│   ├── test-mock-harness.js           #   70 mock PowerPoint run context
│   ├── test-driver-integration.js     # 109 mock driver + radius-core integration
│   ├── test-features.js               #  95 feature behavior (added in v1.3 cleanup)
│   └── README.md
├── dist/                              # build output (gitignored)
│   └── RadiusInPpt.app
├── AGENTS.md                          # required reading: three-layer arch + Mac LTSC pitfalls + §1.0 commit/push rules
├── LOG.md                             # main log (status / done / todo / bugs / plans)
├── README.md                          # this file (zh)
├── README.en.md                       # this file (en)
├── CONTRIBUTING.md                    # zh
├── CONTRIBUTING.en.md                 # en
├── .github/
│   └── SECURITY.md                    # zh
│   └── SECURITY.en.md                 # en
├── changelogs/                        # per-version changelogs
│   ├── v1.0.md
│   ├── v1.1.md
│   ├── v1.2.md                        # v1.2.0 → v1.2.7 main features + early hotfixes
│   └── v1.3.md                        # v1.2.8 → v1.3.0 all hotfixes + v1.3 cleanup + new features
├── plans/
│   └── feature-roadmap.md             # v1.1+ roadmap
├── LICENSE                            # MIT
└── package.json                       # npm test runs 3 test files
```

## Three-Layer Architecture (v1.2 refactor)

```
dialog.js (UI layer)          event binding / rendering / toast / debug log
       │
       ▼
radius-core.js (impl)        10 driver-version functions: writeRadius / readLockState /
                              writeLockState / reapplyLock / applyLayout /
                              syncLayoutChildrenR / pickupFromSelection /
                              applyPickedToSelection / applyLayoutPure / ...
                              zero Office.js calls → 100% unit-testable
       │
       ▼
ppt-driver.js (interaction)  16 methods: load / sync / selectedShapes / activeSlide /
                              slideShapes / shapeId / size / box / isRoundRect /
                              adjFraction / loadAdjValue / setBox / setAdjFraction /
                              addTag / deleteTag / readTag
                              zero business logic → 100% unit-testable
       │
       ▼
Office.js + PowerPoint (Mac LTSC 16.111)
```

- **driver knows no business concepts** (no `LOCK_TAG_KEY` / `LAYOUT_PARENT_TAG_KEY`, no idea what "strict" means)
- **radius-core never imports Office.js** (all shape read/write/load/sync goes through driver)
- **dialog.js is a porter** (`onClick → open driver → call feature → render result`)

## Install in PowerPoint (dev mode)

The manifest references `http://localhost:3000`, so you need to run the local repo first.

### 1. Start the local HTTP server

```bash
# at the project root
npm start
# output e.g.:
#   [serve] HTTP listening on http://127.0.0.1:3000
```

Or just double-click `R 角调整.app` — it auto-starts the server and registers the add-in.

### 2. Register the manifest with PowerPoint

**Official Mac sideload method** ([Microsoft docs](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)):

```bash
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
cp manifest.xml "$WEF/manifest.xml"
```

> `R 角调整.app` does this automatically on launch.

### 3. Load into PowerPoint

1. **Fully quit PowerPoint** (`Cmd+Q` — don't just close the window)
2. **Reopen** PowerPoint
3. The **"主页"** tab → **"加载项"** button
4. Pick **"R 角调整"** in the popup
5. The **"R 角调整"** tab appears in the ribbon

## Usage

### Basic R adjustment

1. **Select** one or more rounded rectangles (multi-select with ⌘)
2. The **"R 角调整"** tab → **"调整 R 角"** button
3. The Dialog opens:
   - Top status card: current selection info + current R (for multi-select, shows min~max)
   - Type a cm value in the input, e.g. `0.3` → click **"应用 R 角"** or press Enter
   - Click **"锁定 R 角"**: locks R to absolute cm on all selected roundRects; click again → all unlock
   - "重新应用锁定 (current selection)": after resizing, use this to re-apply the stored absolute R

### v1.2 Layout Mode

1. **Draw your nesting**: 1 large rounded rectangle (parent) + N small ones (children), positions / sizes arbitrary
2. **Select all** (parent + all children) → **"R 角调整"** tab → **"调整 R 角"**
3. Pick 1 as **parent** (click the parent row in the shape list — "建布局" button appears)
4. Fill **rows × cols** (e.g. `2 × 3`) → pick **child list** → **"建布局"**
5. Drag the **padding / gutter** sliders → children distribute evenly in real time
6. Toggle **R coupling** (off / same / subtract):
   - `off`: child R doesn't move
   - `same`: child R = parent R
   - `subtract` (default): child R = `max(0, parentR − padding)` (natural coupling via padding formula)
7. Change parent R → all child R auto-update per formula

### Lock semantics

v1.1 split "lock" into two independent switches:

| Switch | Behavior |
| --- | --- |
| **Fixed R by value** (button) | write fixed value (cm) → resizing in PPT re-computes to fixed value; dragging the R slider is treated as an intentional edit (updates fixed value) |
| **Anti-misclick / strict** (toggle) | rejects task pane edits + reverses in-PPT R-slider drags to current value; when turned on, auto-locks (uses current R); when turned off, only removes the strict tag, keeps the fixed value |

Common to both:
- Resize the shape → `adjustments[0]` ratio re-computes to match the new short side
- Auto re-apply on selection change (lock monitor 50ms polling detects drag)
- Lock info stored on each shape's `shape.tags` (OOXML `<p:tagLst>`), survives across devices

### Style Brush strict bidirectional (v1.3)

When **"Apply strict-lock state"** is checked, the **source's strict state overrides all targets (bidirectionally)**:
- Source strict=true → add strict tag on all targets (**after** writing R, to avoid writeRadius being blocked)
- Source strict=false → delete strict tag on all targets (**before** writing R, to let writeRadius through)

When unchecked, behavior is unchanged: if any target has strict enabled → the whole style-brush is rejected (anti-misclick invariant).

## Key Technical Points

### Unit conversion

PowerPoint's internal unit is EMU (English Metric Units):
- `1 cm = 360000 EMU`
- **In OOXML** `adjustments[0] ∈ [0, 50000]` (0%~50% of short side)
- **In Mac LTSC Office.js** `adjustments[0] ∈ [0, 1]` (OOXML value ÷ 50000; task pane / dialog contexts both return 0~1)

Let `shortSide = min(width, height)` (EMU):
- **Read**: `radiusCm = adjustments[0] × shortSide / 360000`
- **Write**: `adjustments[0] = clamp(radiusCm × 360000 / shortSide, 0, 0.5)`

> ⚠️ **Don't `Math.round` on write** — `round(0.067) = 0` would truncate every non-integer to 0. `adjustments.set(0, newVal)` accepts 0~1 decimals.

### Why HTTP, not HTTPS

Office Add-ins require HTTPS for production, but `localhost` / `127.0.0.1` is exempt — HTTP is allowed. So we run a plain HTTP server locally — **no cert needed**.

### Mac LTSC Office.js gotchas

See [`AGENTS.md`](./AGENTS.md) §4 for the full list. Key points:
- `shape.adjustments.get(0).value` returns **0~1 decimal**, not OOXML 0~50000
- Must use **collection-level load** (`sel.load('items/...')`); per-shape load doesn't work
- After `set + sync`, must **fresh get(0)** to read new value (old proxy is snapshot-style)
- In Mac LTSC task pane, `customProperties` / `customXmlParts` are unavailable — **only `shape.tags`** works for persistence

## Testing

```bash
cd /Users/ma/Documents/minimax/radius_in_ppt
npm test                                            # run all 4 test files (210 tests)
node test/test-radius-core.js                       # algorithm only (115)
node test/test-mock-harness.js                      # mock harness only (70)
node test/test-driver-integration.js                # driver integration only (109)
node test/test-features.js                          # feature behavior only (95)
```

**Smoke test (in-PPT)**: Task pane → click the "🧪 Driver 烟囱测试" button → 14/14 passing = driver verified.

**Future features go through unit tests**: from v1.3.0, new features no longer require in-PPT testing — trust `npm test` + code review.

## Compatibility

- **PowerPoint for Mac** 2019 / 2021 / 2024 (CustomTab needs Office.js 1.4+)
- **PowerPoint for Windows** 2019+ (uses `localhost:3000` too)
- **PowerPoint for Web** — CustomTab not supported on web, falls back to task pane

## FAQ

**Q: Double-clicking the `.app` shows "cannot verify developer"**
A: Gatekeeper blocks unsigned apps. Right-click the `.app` → Open → in the prompt, click "Open" again (one-time).

**Q: 主页 → 加载项 doesn't show R 角调整**
A: Check:
1. Document is saved to disk (macOS Office Add-ins don't load for unsaved documents)
2. Path `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml` exists
3. Fully quit PowerPoint (`Cmd+Q`) and reopen
4. Show the "开发人员加载项" tab in PowerPoint (if hidden)

**Q: Edits to `src/` don't take effect**
A: Close the Dialog and reopen it.
   If you edited `manifest.xml`, you need to remove and re-add the add-in in PowerPoint.

**Q: How to stop the background server**
A: Terminal: `lsof -ti tcp:3000 | xargs kill`

**Q: How to build + deploy to wef after code changes**
A: `bash tools/build-and-deploy.sh <version> "<commit msg>"` — one-shot (bump version + build + deploy + commit + optional push)

## License

[MIT](./LICENSE)
