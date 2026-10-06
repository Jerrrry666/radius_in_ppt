# RadiusInPpt — Native Mac PowerPoint add-in experiment

Branch: `codex/ribbon-vba-mac`. Target: Office LTSC Standard for Mac 2021.

This branch delivers a `.ppam` loaded by PowerPoint, with radius input, cm/% units, apply, read and presets directly on the ribbon. Daily use requires opening PowerPoint only. [Microsoft documents VBA add-ins and Ribbon XML support on Mac](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac).

**Prototype only: Mac loading, VBA compilation and persistence across PowerPoint restarts have not been validated. The full existing feature set has not yet been migrated.**

## One-time installation

1. Extract `dist/RadiusInPptNative-mac.zip`.
2. Keep the `.ppam` in a permanent location. The included `Install-RadiusInPptNative.command` can prepare a stable copy; manual copying also works.
3. In PowerPoint, use Tools → PowerPoint Add-ins to add the `.ppam`, keep it selected and allow this add-in's macros when prompted.
4. Check for the “R角调整 · Native” ribbon tab, then fully quit PowerPoint with Cmd+Q and reopen to check that the tab remains.

The target is loading with PowerPoint after installation. The helper only copies the file and reveals it in Finder; PowerPoint registration still requires step 3. It is never needed during daily use. Do not register from a build directory that will be rebuilt. Before updating, unload/remove the old entry, replace the file and add it again. See [installation details](native/INSTALL.txt).

## Prototype scope

Sources implement native ribbon input, selection reading, presets, batch radius, write protection and group transactions. Live fixed-radius monitoring, style brush, layout, custom presets and history remain unmigrated. Write protection currently blocks this add-in's edits; it does not undo direct manipulation of PowerPoint handles. Layout-tagged shapes and individually selected group children are rejected. See [design and limitations](plans/ribbon-vba-mac.md).

## Build and verify

```sh
bash tools/build-app.sh
npm test
python3 -m venv .venv-native
.venv-native/bin/pip install -r native/requirements-test.txt
.venv-native/bin/python test/test-native-package.py
```

The existing build entry point now produces a PPAM and an installation ZIP. The ZIP contains only the add-in, a one-time file-preparation helper and instructions. Running the installed plugin requires neither Node nor Python.

The original Office.js code remains as a migration reference: [legacy documentation](README.taskpane.en.md). To build the comparison app explicitly, use `bash tools/build-app.sh --legacy-taskpane`. Legacy app/DMG/wef deployment is not the native installation workflow.

[中文版](README.md) · [Changelog](changelogs/v1.3.md)
