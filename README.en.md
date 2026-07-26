# RadiusInPpt — macOS PowerPoint Add-in

> 🌍 **[中文版](./README.md)**

## About this project

There are already several mature rounded-rectangle add-ins for **Windows PowerPoint**, but on **macOS PowerPoint** there's been no comparable tool.

So I built this one myself — it's still relatively primitive and actively iterating.

**Tested on**: macOS PowerPoint **version 16.111.1 (26071913)** only. Office.js add-ins are cross-platform by design, so it should also work on **Windows** in theory — but **not yet tested on Windows**. Windows users: please try it and report compatibility feedback.

If you also work on rounded rectangles in PowerPoint on Mac, feel free to try it and submit feedback — any thoughts are welcome 🙏

## 📌 Latest update

v1.2 finished "nested rounded-rectangle distribution." v1.3 polishes the layout / style-brush details, completes the dialog.js / radius-core full driver migration, and finishes Step 3-4 migration.

**Layout mode refinements**:
- **Row/column coupled slider** — one slider is enough; columns = children ÷ rows (rows × cols = N strictly, no empty slots)
- **Row values are now a discrete list** — positive factors of N (datalist tick hints), e.g. N=4 → [1, 2, 4], no "3×2=6 → 5 empty slots" cases
- **Padding / gutter Photoshop-style chain link** — chain icon centered between two rows; when active, gutter = padding
- **Gutter disabled when chained** — full greyed-out + non-interactive (eliminates the "gutter modified → chain reverts → shape not reverted" race)

**Style brush strict bidirectional override**:
- **"Apply strict-lock state" = bidirectional override** — source strict=true → target strict=true; source strict=false → target strict=false
- Order matters: source=true → **write R first, then add strict** (avoids writeRadius being blocked); source=false → **delete strict first, then write R** (lets writeRadius through)

**Architecture cleanup**:
- **dialog.js / radius-core fully driver-ified** — 8 driver-version functions replace scattered ctxShape operations
- **Step 3-4 migration complete** — layout tag read/write + pipette go through driver
- **3 long-standing bugs fixed** — style-brush pick-up not applying / layout R-coupling writing only 2 of 4 children / lockMonitor `GeneralException`
- **Bilingual UI** — auto-detects system language (zh / en) for ribbon tab, task pane, and launcher dialogs
- **Tests 210/0** — 95 features + 115 radius-core

Full changelog: [`changelogs/v1.3.md`](./changelogs/v1.3.md)　·　History: [`changelogs/v1.2.md`](./changelogs/v1.2.md)　·　Main log: [`LOG.md`](./LOG.md)

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
- ✅ v1.2 Layout Mode: 1 parent + N children → drag a slider → real-time even distribution + R-coupling
- ✅ v1.3 Style brush: pick the R from one shape, paint to others; optional "apply strict-lock state" bidirectional override

## Features

| Feature | Description |
| --- | --- |
| Custom ribbon tab | "RadiusInPpt" tab + "Adjust R-corner" button (iSlide-style position) |
| Task pane | 360×560 sidebar, modeless, doesn't block view |
| cm / percentage input | `cm` ↔ `%` toggle; % is interpreted as fraction of short side; enter + "Apply R" or Enter key |
| **v1.2** Layout mode | 1 parent + N children, rows×cols grid + padding/gutter sliders + R coupling |
| **v1.2** R coupling | child R = `max(0, parentR − padding)`; off / same / subtract modes |
| **v1.2** Nested state persistence | parent stores JSON + children store parentShapeId as bidirectional tags; travels with the .pptx |
| **v1.3** Row/column coupled slider | one slider, columns = children ÷ rows (rows × cols = N strictly) |
| **v1.3** Row discrete factor list | valid row values = positive factors of N ([1, 2, 4] / [1, 2, 3, 6] / prime [1, N]), no empty slots |
| **v1.3** Padding/gutter chain link | Photoshop-style chain icon; when active, gutter = padding; gutter fully disabled when chained (avoids race conditions) |
| **v1.3** Style brush strict bidirectional | checkbox toggles bidirectional override: source strict=true → target strict=true; source strict=false → target strict=false |
| Fix R by value | button on/off; when on, resizing in PPT re-computes to fixed value |
| Anti-misclick (strict) | independent toggle; when on, uses current R as fixed value; rejects task pane edits + reverses R-slider drags |
| R preset library | 5 user-editable presets, name + value, one-click apply |
| R style brush | pick from 1 roundRect's R, paint to other targets; optional "apply strict-lock state" bidirectional override |
| Multi-select | acts on all selected rounded rectangles; non-roundRects are skipped with a notice |
| Shape list | live display of each selected shape's current R (syncs with in-PPT edits) |
| Auto-reapply on lock | setInterval polling, distinguishes "resize drag" from "R-slider drag" |
| Persistence | lock info stored on each shape's `shape.tags` (OOXML `<p:tagLst>`), survives across devices |

## Usage

Distributed as a macOS `.app` (~405 KB), double-click to launch:

1. Double-click `RadiusInPpt.app`
2. A prompt appears → choose "Quit and reopen PowerPoint"
3. Reopen PowerPoint → **"RadiusInPpt"** tab appears in the ribbon
4. Click **"Adjust R-corner"** in the tab → task pane opens on the right
5. Select a rounded rectangle → type `0.3` (cm) or `10` (%) in the task pane → Apply / Lock / Strict
6. **v1.2 Layout Mode**: select 1 parent + N children → Group → "Build layout" → drag sliders to distribute in real time

> After that, just double-click the `.app` every time you need it (server runs in background, manifest is persistent).
> Note: if you changed code, fully quit PowerPoint (`Cmd+Q`) and reopen for the task pane to pick up new code.

### Windows install

A **Windows version** is also available from [GitHub Releases](https://github.com/Jerrrry666/radius_in_ppt/releases/tag/v1.3) as `RadiusInPpt-win.zip` (~95 KB). **Not yet tested on Windows** — please report issues.

1. Download `RadiusInPpt-win.zip` from the Releases page
2. Extract to any folder
3. Make sure [Node.js 18+](https://nodejs.org/) is installed (the `.bat` launcher auto-detects it)
4. Double-click `RadiusInPpt.bat` → popup says "fully quit PowerPoint and reopen"
5. Open PowerPoint → fully quit (File → Exit) → reopen
6. **"RadiusInPpt"** tab appears in the ribbon

> After that, just double-click `RadiusInPpt.bat` to start the server each time.
> Log location: `%TEMP%\radius_in_ppt.log` (attach when reporting issues).

## FAQ

**Q: Double-clicking the `.app` shows "cannot verify developer"**
A: Gatekeeper blocks unsigned apps. Right-click the `.app` → Open → in the prompt, click "Open" again (one-time).

**Q: Home → Add-ins doesn't show RadiusInPpt**
A: Check:
1. Document is saved to disk (macOS Office Add-ins don't load for unsaved documents)
2. Path `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml` exists
3. Fully quit PowerPoint (`Cmd+Q`) and reopen
4. Show the "Developer Add-ins" tab in PowerPoint (if hidden)

**Q: Edits to `src/` don't take effect**
A: Close the Dialog and reopen it.
   If you edited `manifest.xml`, you need to remove and re-add the add-in in PowerPoint.

**Q: How to stop the background server**
A: Terminal: `lsof -ti tcp:3000 | xargs kill`

## License

[MIT](./LICENSE)
