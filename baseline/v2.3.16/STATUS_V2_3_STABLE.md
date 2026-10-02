# Mimaki Imposer v2.3.16 - stable TIFF baseline

Status date: 2026-08-12
Stable version: v2.3.16
Branching model: local-only development, no GitHub repo yet
Frozen snapshot folder: `STABLE_v2_3_16_2026-08-12`

This version is the approved production fallback for the current TIFF-based imposition workflow.

## Locked Workflow Scope

- Template path is strict: `S:\printstudio\zakazky\.VelFo\mimaki_sablonatlac.cdr`.
- Product/template layers are read from names ending with `SLOT`.
- Slot objects are only shapes with outline spot color `CutContour1`; other objects on those layers are ignored as slots.
- Normal `Print Data` goes through the 600 dpi CMYK TIFF temp-file workflow.
- `RDG_WHITE` spot objects are detected by spot color only, bypass TIFF, and remain vector spot data on `WHITE`.
- Every slot uses its own template geometry plus 3 mm bleed on each side.
- Card F3 helper is card-only and cycles 85 x 54 / 54 x 85 mm page setup with boundary guides.
- P/Z workflows are card-only; non-card products stay single-sided.
- Output CDR saves automatically beside the source, with fallback to the configured default folder.
- Production PDF export hides non-printable template/slot layers and exports `_W` PDFs only when WHITE data exists.
- `KOPIE` mode defaults piece count to the selected template capacity.
- Open Template macro uses classic behavior and opens/activates the real template document.

## Stability Rule

Do not change the v2.3.16 TIFF workflow mechanics while developing new features.

Future "bez TIFF / zachovat vektor" work must be built as a separate workflow path, not by replacing or weakening the stable TIFF path. If new development breaks output, restore from `STABLE_v2_3_16_2026-08-12`.

## Snapshot Files

- `Mimaki_Template_Imposer_v2_3.bas`
- `Mimaki_Template_Imposer_v2_3_CODE_ONLY.txt`
- `frmMimakiImposerPanel.frm`
- `frmMimakiImposerPanel_CODE_ONLY.txt`
- `AKTUALNY_MIMAKI_IMPORT_V2_3.txt`

## Snapshot SHA256

- `Mimaki_Template_Imposer_v2_3.bas`: `7677D05FD7F97DB88F83C6EE0F6559EACCF715387EAECFDE6FC573E74CB795DB`
- `Mimaki_Template_Imposer_v2_3_CODE_ONLY.txt`: `7677D05FD7F97DB88F83C6EE0F6559EACCF715387EAECFDE6FC573E74CB795DB`
- `frmMimakiImposerPanel.frm`: `B16EF443D0F8F530DEA3F39DDA72AAD0C944BF4605E7748E5282CB4F7A782F07`
- `frmMimakiImposerPanel_CODE_ONLY.txt`: `06506B081DBBE6A13814F650E45E59F5A9ED86BBFACD2F3220A476544ACF77C9`
- `AKTUALNY_MIMAKI_IMPORT_V2_3.txt`: `2700E8383D5742EAE8598FCC1D8A528C10E9C438E0DE684D743A66D5EDC00879`
