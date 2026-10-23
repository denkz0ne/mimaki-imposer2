# VBA Import Order

For CorelDRAW 2026, import the v2.4 standard modules and form into a new VBA project, then compile before first use:

1. `Mimaki_Template_Imposer_v2_4.bas` (core and legacy TIFF path)
2. `modMimakiPlacementVector.bas` (isolated vector path called by the core)
3. `frmMimakiImposerPanel.frm` (panel)

The panel is a native designer form. Keep `frmMimakiImposerPanel.frx` next to the
FRM during import: it contains all 50 controls, their editable positions, fonts,
colors and embedded icons. Import the FRM, not the TXT. No separate icon folder
or SDK is needed. Existing toolbar macro names and core modules are unchanged.

To change layout, open `frmMimakiImposerPanel` in the VBA editor and choose View
Object (Shift+F7). Move or resize controls and save the GMS. Keep `(Name)` values
unchanged. Runtime presentation updates change visibility, captions and selection
colors, but never replace the designer's coordinates or dimensions. The copies
and page-range fields intentionally occupy the same position; choose them by
name in the Properties window when editing.

`frmMimakiImposerPanel_CODE_ONLY.txt` has no header and contains only the form's
event code. It updates a form that already has these designer controls; pasting
it into an empty UserForm does not create the layout or icons.

Developer regeneration uses `tools/build-panel.ps1` with the installed CorelDRAW
2026 VBA editor. The script creates and unloads a separate scratch GMS and exports
a genuine FRM/FRX pair. `tools/test-panel.ps1` performs a native import, checks all
control coordinates, and runs UI event/geometry/icon tests without imposition.
Install the verified form manually using the import steps above. The workstation
deployment was performed separately with a local backup and exact comparisons
of the existing core/vector modules.

Keep v2.3.16 installed as a separate stable GMS/project until v2.4 passes the CorelDRAW validation matrix. Do not replace the stable module in-place.

The vector module intentionally calls a small set of public Corel/VBA helpers in the core module. The established TIFF call path remains routed directly to `PlaceSourceRangeIntoSlot`; selecting vector preservation is the only condition that enters the new module.
