# Mimaki Imposer v2.4 for CorelDRAW 2026 — Specification

## Status

- Stage: approved; implementation in progress
- Target host: CorelDRAW Graphics Suite 2026, API v27.2
- Production baseline: Mimaki Imposer v2.3.16 (2026-08-12)
- Repository: https://github.com/denkz0ne/mimaki-imposer2
- This document is not implementation approval. The frozen baseline must remain recoverable.

## Goal

Deliver a maintainable and measurably efficient CorelDRAW 2026 version of Mimaki Imposer while preserving the production output and behavior of v2.3.16. Add a separate vector-preserving placement path that keeps source artwork editable and avoids TIFF rasterization, without replacing or weakening the proven TIFF path.

## Target and API policy

- Target the documented CorelDRAW 2026 v27.2 object model.
- Keep the implementation as an in-process Corel VBA macro/UserForm unless a documented and demonstrated limitation requires a different host technology.
- Prefer direct, documented, typed object-model calls for the 2026-only build. Retain late-bound calls only where the object model or export interfaces require them and document why.
- Use `ShapeRange.ConvertToBitmapEx` for raster conversion; do not call the obsolete `ConvertToBitmap` first.
- Use `FindShapes`/CQL only where it preserves the strict template/spot-color semantics; do not assume a CQL query can distinguish spot color unless verified against CorelDRAW 2026.
- Keep expensive source/template scans bounded to the relevant document, page, and layer. Avoid selection-dependent operations when a shape range or explicit layer target can be used.
- Do not claim performance gains without timing repeatable representative jobs in CorelDRAW 2026.

## Frozen production behavior

The following behavior is a compatibility contract for the default TIFF workflow:

1. The template is opened from the configured canonical template path and treated as read-only by the imposer.
2. Product layouts are template layers whose names end in `SLOT`. A slot is only an object with outline spot color `CutContour1`; other objects are not slots.
3. Slot geometry comes from the selected template layer. No product-specific fixed slot or artwork dimensions may be assumed.
4. Artwork is fitted to each slot's own dimensions plus 3 mm bleed on every side.
5. Print Data is rasterized to CMYK TIFF at 600 dpi, imported, fitted, and placed using the established orientation/crop behavior.
6. Objects whose spot color is `RDG_WHITE` bypass TIFF and remain editable vector objects with the original spot color on the WHITE output layer. Detection is based on spot color, independent of source layer name.
7. The output CDR is saved automatically beside the source document. If that location is unavailable, use the configured default output folder.
8. The PDF export hides non-printing template/slot layers. Export the normal PDF for each output page; export a second `_W` PDF for a page only when that page contains WHITE data. Suppress success dialogs.
9. F3 is restricted to cards. It alternates page size between 85 x 54 mm and 54 x 85 mm and replaces the prior card boundary guides at the page corners. It does not alter artwork.
10. P/Z/PZZ/PPZ card workflows remain card-only. Other products are single-sided and use page-copy or page-layout modes.
11. In the imposed CDR, the template slot layer remains above Print Data, with other layer order and visibility governed by the existing workflow.
12. The panel continues to refresh product/slot information when the selected layout changes and reports both slot size and slot-plus-bleed artwork size.
13. Source artwork remains unchanged after a run. Any required temporary document state is undone or removed, and Corel optimization/command groups are restored and closed even after an error.
14. No modal success dialog is shown after imposition or successful CDR/PDF saves.

## Separate vector-preserving workflow

Add a per-product `Zachovať vektor` option. Its behavior is independent from the default TIFF workflow:

- When enabled, do not rasterize Print Data and do not create TIFF intermediates.
- Preserve the original vector artwork and its layer structure while placing it in the slots. Rotation, scale, alignment, and bleed-area fitting must match the established TIFF placement geometry.
- Keep `RDG_WHITE` objects as vectors with their spot color and route them to WHITE under the same color-only detection rule.
- Do not modify or save changes into the source document or read-only template. Temporary operations must be reversible or fully cleaned up.
- If an input cannot be placed faithfully as vectors, stop with a useful error before producing a misleading partial output; do not silently fall back to TIFF.
- The unchecked option must execute the frozen TIFF path with equivalent output to v2.3.16.

## Architecture constraints

- Keep the v2.3.16 source snapshot immutable and identifiable by hashes.
- Develop the new implementation separately from the stable snapshot. The TIFF and vector-preserving paths may share validated template parsing, slot geometry, orientation, and export utilities, but must have distinct artwork preparation/placement processors.
- Keep the UI as a thin controller: read options, validate, invoke the workflow, report progress/errors.
- Use explicit cleanup/error-exit paths so every successful `BeginCommandGroup` is matched with `EndCommandGroup`, and optimization/redraw state is restored to its prior value.
- Avoid changing F3, duplex card logic, template detection, save destinations, or export semantics as collateral effects of the vector workflow.

## Acceptance criteria

- The full behavior contract above is verified against representative PVC card, pen, USB, fragrance, and other single-sided SLOT layouts.
- For a fixed source/template/settings fixture, the default TIFF workflow produces the same slot count, output page count, per-slot geometry, orientation, layer order, spot-color separation, save location, and PDF naming/WHITE behavior as v2.3.16.
- Vector-preserving mode produces editable vector objects and preserves source layer structure; no TIFF is generated, and WHITE remains a spot vector.
- A source document reopened after processing has the same page content, selection, and relevant document state as before the run.
- Errors during import, placement, save, and PDF export leave no orphan temporary objects/files and do not leave CorelDRAW in optimization or command-group state.
- Performance is compared against v2.3.16 on the same CorelDRAW 2026 workstation and representative jobs. Record timings and environment; only report a speedup when measured.
- The project documents the supported CorelDRAW 2026 version, install/import procedure, known limitations, and rollback path.

## GitHub delivery model

- Preserve v2.3.16 as the reviewable stable reference and add a version tag after its files are committed.
- Develop v2.4 on a separate branch; review changes through pull requests before updating the stable line.
- Keep source exports (`.bas`, `.frm`), import-ready text copies, release notes, and validation fixtures/results versioned. Do not commit customer artwork or production job files.
- Keep generated/runtime GMS files separate from source; only publish one after it has been built and verified in CorelDRAW 2026.
