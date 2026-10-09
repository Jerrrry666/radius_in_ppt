# RadiusInPpt v1.4.0 — Native Mac PowerPoint add-in

Default branch: `main`, including the native v1.4.0 add-in. Target: Office LTSC Standard for Mac 2021.

The project delivers a `.ppam` loaded by PowerPoint, with radius input, cm/% units, apply, read, presets and parent/child relationships directly on the ribbon. Daily use requires opening PowerPoint only. [Microsoft documents VBA add-ins and Ribbon XML support on Mac](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac).

**Loading, ribbon edits and automatic loading after a full quit were verified on this Mac with PowerPoint 16.113.3 on 2026-10-07. This remains a migration prototype; the complete feature set and the 16.111 target build still need separate validation.**

## One-time installation

1. Extract `dist/RadiusInPptNative-mac.zip`.
2. Keep the `.ppam` in a permanent location. The included `Install-RadiusInPptNative.command` can prepare a stable copy; manual copying also works.
3. In PowerPoint, use Tools → PowerPoint Add-ins to add the `.ppam`, keep it selected and allow this add-in's macros when prompted.
4. Check for the “R角调整 · Native” ribbon tab, then fully quit PowerPoint with Cmd+Q and reopen to check that the tab remains.

The target is loading with PowerPoint after installation. The helper only copies the file and reveals it in Finder; PowerPoint registration still requires step 3. It is never needed during daily use. Do not register from a build directory that will be rebuilt. To update the same path, save your documents and fully quit PowerPoint before replacing the file, then reopen and verify loading. When changing paths, remove the old entry and add the new path. The default installation directory is `~/Library/Application Support/RadiusInPptNative`. See [installation details](native/INSTALL.txt).

## Current scope

Native ribbon input, selection reading, presets, batch radius, write protection, parent/child grid layout and radius links are available. Up/down arrows apply ±0.1 in the current unit immediately. Protection status and counts refresh with the selection; protected selections disable radius writes. All action icons are embedded PNGs. Live fixed-radius monitoring, style brush, custom presets and history remain unmigrated. Write protection blocks this add-in's edits; it does not undo direct manipulation of PowerPoint handles. Legacy layout-tagged shapes support explicit radius/protection edits while preserving metadata. Radius/protection edits on individually selected group children require selecting the complete top-level group. See [design and limitations](plans/ribbon-vba-mac.md).

The radius field uses the same width and arrow spacing as the layout padding/gap fields, with separate radius/unit labels to reduce unused space. Mac displays custom up/down arrows side by side to the right of the field.
This radius control update was verified on PowerPoint16.113.4/26100421; see its [acceptance record](test/native-host-validation-radius-controls-20261007.json).

## Parent/child relationships

1. Select one rounded rectangle and click “设为父对象” (set parent). The ribbon shows its name while binding is pending.
2. Select the child rounded rectangles and click “绑定子对象” (bind children). A group containing the intended children can be selected in one operation; non-rounded members remain untouched. Set an existing native parent again to append children. Changing slides or documents cancels the pending parent.
3. “查看关系” (view relationships) lists slide-local G01, G02, etc., with parent and child names. Click a member to locate it, or select all children/the whole relationship. Selection status shows group codes, roles and the number of unbound objects.
4. “编号预览” (numbered preview) opens a disposable copy of the current slide with blue parent badges and brown numbered child badges. Click again to close it and return to the source; source saves and exports contain no badges. The preview is a static snapshot, with add-in edits disabled. Closing it discards any changes to the copy.
5. “解除关系” (detach) removes selected children or, after confirmation, the whole relationship. Geometry, radius, protection and unrelated tags remain intact.

Relationships persist in the document and apply to rounded rectangles on the same slide. Child numbers follow member order when bound; removing a child retains the other numbers. Protected shapes can still be bound or detached. Legacy layouts can be viewed and explicitly detached before rebuilding. Reparenting requires explicit detachment. Copying tagged members can create duplicate parents or child numbers; relationship operations stop on such conflicts until the tags are repaired.

## Automatic layout and radius links

After binding, select the parent or a child and configure rows, columns, padding and gap under “自动布局与R角联动”. “应用布局” distributes children in equal grid cells in child-number order. Changing rows or columns adjusts the other dimension. Padding and gap are in centimeters and stay constant when the parent is resized.

Child radius modes are same as parent, parent radius minus padding (minimum zero), and off. Each child is clamped to half its own shorter side. Off preserves each child's adjustment fraction while geometry can still follow the parent.

This control update was tested on PowerPoint 16.113.4/26100421; its [acceptance record](test/native-host-validation-layout-controls-20261007.json) is separate from the earlier 16.113.3 validation.

The compact parameter columns group rows/columns/radius mode and padding/gap/status, with actions on the right. Arrow buttons change rows/columns by 1 and spacing by 0.1cm, bounded at zero. They stage parameters until “应用布局” commits them. Mac displays custom up/down buttons side by side beside each field.

The first layout application enables automatic linking. Add-in radius edits to the parent update children immediately. After directly moving/resizing the parent or dragging its yellow handle, click a blank area or change selection to synchronize; saving also synchronizes. Updates do not run on every drag frame. Disable “自动联动” to stop following parent edits; manual application remains available. Settings survive saving and reopening.

A protected child rejects the entire layout/linked-radius batch and disables the related parent controls. Insufficient space rejects all writes. Scaled nested groups use safe ungroup/regroup transactions preserving names, other tags, leaf IDs and hierarchy. Rotated or flipped members/groups are currently unsupported for layout and linked radius.

## Build and verify

For a quick logic check, click “R角调整 · Native → 快速自检” in PowerPoint. It creates its own temporary presentation, checks algorithms and real radius/protection/group/relationship/layout operations, then closes the test document and returns to the source. Results show actual counts, errors and elapsed time. It does not save a test PPTX; close any numbered preview first.

See the [quick-test record](test/native-host-validation-quickcheck-20261009.json) for actual counts, timing, package hashes and host coverage. Saved files, all control events and recovery from host write failures use separate validation protocols.

For project checks, double-click [Run-Tests.command](tools/Run-Tests.command) or use `npm run test:quick`. See the [test checklist and lessons](test/README.md) for coverage. Actual group-child selection, control events, save/reopen and host write-failure recovery have separate validation protocols; a passing quick report does not establish all of them.

```sh
npm run build
python3 -m venv .venv-native
.venv-native/bin/pip install -r native/requirements-test.txt
npm test                     # Native format, installation and saved-state reader
npm run test:legacy          # Office.js migration reference
npm run test:all             # Both suites
npm run test:quick           # One-click project checks; also a Finder .command entry
```

The existing build entry point produces a PPAM and an installation ZIP. The ZIP contains only the add-in, a one-time file-preparation helper and instructions. Running the installed plugin requires neither Node nor Python. `npm test` selects a Python interpreter with the test dependencies, preferring the project's `.venv-native`; these consumer checks do not execute VBA. Actual ribbon operations and saved OOXML results are recorded separately for [relationships](test/native-host-validation-relations-20261007.json) and [layout/radius links](test/native-host-validation-layout-20261007.json).

The original Office.js code remains as a migration reference: [legacy documentation](README.taskpane.en.md). To build the comparison app explicitly, use `bash tools/build-app.sh --legacy-taskpane`. Legacy app/DMG/wef deployment is not the native installation workflow.

The [2026-10-09 main validation record](test/native-host-validation-main-20261009.json) distinguishes completed and pending checks for the current refactor. Updates do not create old-package backups; host test results are recorded before temporary PPTX files are removed.

[中文版](README.md) · [Changelog](changelogs/v1.4.md)
