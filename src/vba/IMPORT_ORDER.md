# VBA Import Order

For CorelDRAW 2026, import the v2.4 standard modules and form into a new VBA project, then compile before first use:

1. `Mimaki_Template_Imposer_v2_4.bas` (core and legacy TIFF path)
2. `modMimakiPlacementVector.bas` (isolated vector path called by the core)
3. `frmMimakiImposerPanel.frm` (panel)

Keep v2.3.16 installed as a separate stable GMS/project until v2.4 passes the CorelDRAW validation matrix. Do not replace the stable module in-place.

The vector module intentionally calls a small set of public Corel/VBA helpers in the core module. The established TIFF call path remains routed directly to `PlaceSourceRangeIntoSlot`; selecting vector preservation is the only condition that enters the new module.
