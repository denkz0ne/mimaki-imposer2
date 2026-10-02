# v2.4 Implementation Progress

- Work branch: work/v2.4-corel-2026
- Stable ref: stable/v2.3.16
- Baseline commit: 38737845f83d8a50242c87c8081c288ddfb21e54
- Frozen source SHA-256 values match STATUS_V2_3_STABLE.md; GitHub blob IDs match local git hash-object --no-filters.
- Connector limitation: no tag-creation tool is exposed, so stable/v2.3.16 is used instead of claiming a v2.3.16 tag.
- The user approved implementation and authorized autonomous work.
- CorelDRAW is not exposed as a native app in the current computer-use app inventory. A real Corel compile/runtime test has not been performed; no GMS build will be claimed.
- Keep unverified runtime/API assumptions open until tested in CorelDRAW 2026 v27.2.

## Work Log

### Task 1: Stable baseline
- Complete: frozen source/docs uploaded unchanged.
- Complete: stable branch and v2.4 work branch created.
- Complete: exact blob hashes checked for all six source files.
- Baseline stays unchanged throughout the v2.4 branch.
- Vector placement extracted to `src/vba/modMimakiPlacementVector.bas`; TIFF branch remains direct to stable processor.
- Header/version metadata corrected to v2.4.0 / 2026-10-02.
- Static Python contract tests added; CorelDRAW 2026 compile/runtime matrix is still pending on the target host.
