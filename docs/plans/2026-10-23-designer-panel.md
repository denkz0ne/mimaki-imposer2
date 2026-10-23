# Editable VBA Panel

## Approved Direction

Implement the approved local HTML preview as a real MSForms UserForm. All controls,
their default coordinates, dimensions, tab order, colors, fonts and icons must live
in the form designer and its exported FRX resources. Users can drag and resize them
in the VBA editor. Keep existing control names, settings and placement entry points.
No SDK, runtime, font or package installation is permitted on this workstation.

Use the v2.4 development branch at d03a98b as the source. The installed GlobalMacros
project contains the same generation, with locally adjusted paths. Export a local
backup before any change to that project. Core and vector placement modules stay
byte-identical in Git and unchanged in the installed project.

## Implementation Plan

- [x] Replace the runtime BuildUi and icon extraction with designer controls.
  Use Segoe UI, square controls, light gray surfaces, green selected states and
  embedded small bitmap icons. Keep VBA source ASCII; localized designer captions
  and pictures are exported through FRX by the installed Corel VBA editor.
- [x] Add a repeatable Windows PowerShell builder using Corel/VBIDE COM. Create an
  isolated build GMS; export a genuine FRM/FRX pair and code-only TXT, then unload
  the build project. The build script must not alter the installed production form.
- [x] Adapt only presentation procedures: metric labels, explanatory text, state
  styling and visibility. Never reset designer coordinates or dimensions at runtime.
  Keep the placement command handler and source selection functions unchanged.
- [x] Verify the native form by import/export round-trip, designer control counts,
  control bounds and names, event compilation and synthetic UI interaction. Do not
  run imposition on customer documents. Run the existing placement contract suite.
- [x] Package the form and instructions. Import the verified panel into the existing
  project after backing up its form, retain the existing core modules and toolbar
  identifiers, and show the native result. Save a reviewable Git branch/PR.

## Review Focus

Check duplicate designer/WithEvents declarations, lost runtime event hookups,
Corel COM errors during import, stale hidden input values, source selection changes,
control overflow at Windows DPI settings, and preserved geometry after manual edits.

## Constraints And Decisions

- UI generation date is 2026-10-23; the caption also retains the core script version.
- CODE_ONLY is for existing designer forms, not a substitute for importing FRM/FRX.
- GMS build files and installed-project backups are local-only and never published.
- Native Windows borders may differ from the HTML preview.

## Verification Record

- CorelDRAW 2026 27.2.0.135; existing installed VBA/Forms engine only.
- Genuine FRM/FRX export and import: 50 controls, all coordinates within 0.25 pt.
- Native probe: embedded pictures, designer movement across initialization/mode
  changes, card/single-sided workflows, copies/range visibility, orientation,
  sides, vector/TIFF processing state and existing-output availability passed.
- Existing Python placement contracts: 8/8 passed; designer Node contract passed.
- Installed in the existing GlobalMacros project after local-only backups.
  Other modules compared byte-for-byte before/after; unchanged. Saved via the
  VBA editor's Save command after confirming it targets GlobalMacros.
- Actual installed panel opened and its accessible control tree showed product,
  slot/bleed dimensions and all expected controls. Screenshot capture of Corel
  timed out; visual pixel inspection remains unverified.
- No production imposition/export test was run for this UI-only change.
- Review found that the initial native probe's QueryClose persisted test defaults
  under the normal preferences namespace. The test harness now binds both the
  scratch form and probe to a disposable namespace before initialization. The
  earlier orientation/side defaults were not captured before that first probe;
  inform the operator to check them. No production imposition was executed.
- The workstation deployment helper is local-only. It is not distributed as an
  installer: on failure it restores the form but cannot guarantee full-project
  rollback. The deliverable uses the normal manual FRM/FRX import procedure.
