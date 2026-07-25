# Contributing

> 🌍 **[中文版](./CONTRIBUTING.md)**

Thanks for considering contributing to **RadiusInPpt** (R 角调整)! 🎉

This is a macOS PowerPoint add-in that lets users set the corner radius of rounded rectangles in **centimeters** or **percentages**.

## Filing Issues

- **Bug reports** — Use [GitHub Issues](https://github.com/Jerrrry666/radius_in_ppt/issues/new) with a short, descriptive title.
- **Feature requests** — Same as above, but prefix the title with `[Feature Request]`.
- **Q&A / discussion** — Use [GitHub Discussions](https://github.com/Jerrrry666/radius_in_ppt/discussions) (if enabled).

Please search existing issues first to avoid duplicates.

## Submitting a PR

### 1. Fork + branch

```bash
# after forking
git clone https://github.com/<your-username>/radius_in_ppt.git
cd radius_in_ppt
git checkout -b feat/your-feature-name
```

### 2. Code conventions

- **Three-layer architecture** — `dialog.js` (UI) → `radius-core.js` (algorithm) → `ppt-driver.js` (Office.js interaction). Think carefully about which layer your change belongs in.
- **driver knows no business concepts** — it must not recognize `LOCK_TAG_KEY`, `LAYOUT_PARENT_TAG_KEY`, or any other domain keys.
- **radius-core must not import Office.js** — all Office calls go through the driver.
- **AGENTS.md** is required reading for AI agents (Mac LTSC Office.js pitfalls). Human contributors are also encouraged to skim §1-§2.

### 3. Run tests

```bash
npm test
```

You should see 210/0 passing (or more as the suite grows). New features must include unit tests.

### 4. Commit messages

- Write in English. **Avoid CJK punctuation** (commit messages get parsed by bash pipelines).
- Format: `<scope>: <what changed>`
  - `feat: add X`
  - `fix: resolve Y`
  - `refactor: Z`
  - `docs: ...`
  - `test: ...`
- One logical change per commit.

### 5. Open the PR

- Clear title describing the change.
- Body should include:
  - **What** — what you changed
  - **Why** — why (which issue / use case)
  - **How to test** — verification steps
  - **Screenshots** — for any UI change
- Link related issues (`Fixes #123`).

## First-time setup (developers)

```bash
git clone https://github.com/Jerrrry666/radius_in_ppt.git
cd radius_in_ppt
npm start   # starts the local HTTP server (localhost:3000)
```

Then register the manifest with PowerPoint:

```bash
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF" && cp manifest.xml "$WEF/manifest.xml"
```

Fully quit PowerPoint (`Cmd+Q`), reopen, and the **"R 角调整"** tab will appear on the ribbon.

## Build

```bash
bash tools/build-app.sh   # produces dist/RadiusInPpt.app
```

## Communication style

- Direct, friendly, on-topic.
- No need for formal pleasantries in issues / PRs.
- Code review is fact-based, not authority-based.

## License

By submitting a PR, you agree to license your contribution under the [MIT License](./LICENSE).
