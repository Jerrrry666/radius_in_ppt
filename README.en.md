<p align="left">
  <a href="README.md"><kbd>简体中文</kbd></a>
  <a href="README.en.md"><kbd><strong>English</strong></kbd></a>
</p>

# Radius in PowerPoint · Native

Native macOS PowerPoint add-in · v1.4.0

Edit rounded corners, manage parent–child relationships, and arrange child shapes in a grid from the PowerPoint ribbon. After installing the `.ppam`, use the **R角调整 · Native** tab. Daily use does not require a separate app or server.

## Features

- **Batch radius editing:** Set a radius in centimeters or as a percentage of each shape’s shorter side. Read the current value, apply a preset, or edit multiple shapes at once. Centimeter values apply equally to every target; percentages are calculated for each shape. The radius is capped at half the shorter side, and 0 is allowed.
- **Write protection:** Protection status and counts update with the selection. If a selected shape is protected, the add-in disables the relevant controls and checks protection again before writing. Protection blocks radius edits made by this add-in.
- **Quick adjustment:** The arrows beside the radius field change the value by 0.1 in the current unit and apply it immediately.
- **Parent–child relationships:** Assign one parent rounded rectangle and bind multiple children on the same slide. View or locate members, detach them, or preview numbered members in a temporary copy that does not modify the source presentation.
- **Grid layout and radius links:** Arrange children in rows and columns with centimeter padding and gap. Child radius can match the parent, subtract the padding, or remain independent. Relationship and layout settings are saved in the presentation.
- **Quick check:** Run algorithm and business checks from the ribbon. Checks use a temporary presentation, which is closed without saving a test file.

## Install

1. Extract `dist/RadiusInPptNative-mac.zip` and keep the `.ppam` in a stable location. The default folder is `~/Library/Application Support/RadiusInPptNative`. Use the included file-preparation script or copy the file manually.
2. In PowerPoint, open **Tools → PowerPoint Add-ins**, add the `.ppam`, and leave it enabled. Allow macros for this add-in if Office prompts you.
3. Confirm that the **R角调整 · Native** tab appears. Quit PowerPoint completely with Cmd+Q, reopen it, and confirm the add-in loads again.

The file-preparation script only copies the file; PowerPoint registration is a separate step. To update the add-in, save your presentations and quit PowerPoint before replacing the installed file, then reopen PowerPoint to verify it. Do not load the add-in from the `dist` folder, which may be rebuilt. See the complete [installation guide](native/INSTALL.txt).

## Use

Select one or more rounded rectangles, enter a centimeter value or percentage in the ribbon, and choose **应用R角** (Apply Radius). You can also read the current value, use a preset, or click an arrow to apply a small adjustment. Percentages use each selected shape’s own shorter side.

To create a layout, select the parent and choose **设为父对象** (Set Parent), then select the children and choose **绑定子对象** (Bind Children). Select a member of the relationship to configure rows, columns, padding, gap, and child-radius mode, then apply the layout. Layout and radius linking are unsupported for rotated or flipped groups and members. If any child is protected, the entire layout or linked-radius write is rejected.

## Compatibility and limits

- The current deliverable is a PowerPoint VBA add-in (`.ppam`), loaded by PowerPoint; it is not a standalone app.
- The target is Office LTSC Standard for Mac 2021. Current host validation includes PowerPoint 16.113.4/26100421. The target build 16.111 still requires separate validation.
- To edit the radius or protection state of shapes inside a group, select the complete top-level group. Parent–child relationship actions support group members.
- Write protection blocks edits made by the add-in. It does not undo direct changes made with PowerPoint’s yellow adjustment handles or by resizing a shape.
- Live fixed-radius monitoring, the style brush, custom preset libraries, and history have not been migrated to the native add-in.

## Build and checks

```sh
npm run build          # Build the PPAM and macOS installation ZIP
npm run test:quick     # Native PPAM format and installation checks
```

The installed add-in does not require Node or Python. See the [test guide](test/README.md) for coverage. The [native design](plans/ribbon-vba-mac.md), [project log](LOG.md), and [v1.4.0 changelog](changelogs/v1.4.md) contain architecture, limitations, and validation records.
