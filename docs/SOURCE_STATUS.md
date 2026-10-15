# Source status

`vba/` is an export of the VBA project of the workbook in use for the 2026/27
school year. It was made on 2026-10-15 with `tools/Export-WorkbookVba.ps1`, which
reads a copy of the workbook with macros disabled. No macros, self-checks, or
templates were run for this snapshot.

## Included sources

| Component group | Source |
| --- | --- |
| 52 standard modules and four classes | Exported from the workbook |
| `frmProgress.frm` and `.frx` | Exported from the workbook; replaces the copy kept from the 2025 repository |
| `Kinder.vba` and `Kinder_Blanks.vba` | Worksheet code from the workbook, export headers removed |
| `ThisWorkbook.vba` | Workbook code from the workbook, export header removed |

Production workbooks, administrative data, templates, generated output, temporary
Office files, internal notes, and task prompts are excluded. The repository is not
a backup of the operational workbook system.

## Changes since the previous snapshot

- The output routing that the previous snapshot described as ahead of the
  workbook is now part of it: `GENERATE_WORD_FOR_NON_JOBCENTER` is `False`,
  non-Jobcenter/Sozialamt records produce only the legacy XLSX form, and the
  record processor starts Word only when DOCX output is needed.
- `ThisWorkbook` code is now included. It declares `IsMacroRunning`, which
  `Archivierung` and `Synchronisation` set while they run; a workbook built from
  the previous snapshot alone did not compile.
- `frmProgress` now comes from the workbook. Its code and controls match the
  earlier copy.
- The remaining differences are spelling only. In 16 modules and one class the VBA
  editor changed the case of identifiers (for example `NumberFormat` to
  `numberFormat`), and the worksheet files keep the indentation of blank lines.
  The logic is unchanged.

## Not included

A separate AdminSync extension that adds missing records from the administrative
workbook to `Kinder`, including restoration from `Archiv`, is part of neither the
workbook nor this snapshot. The included AdminSync synchronizes existing records.

## Verification limits

The export shows which code the workbook contains, not that every workflow runs
correctly in a freshly integrated workbook. In particular:

- UTF-8 source text must survive the target VBA project's encoding.
- Worksheet and workbook code must be attached to the correct document modules.
- Templates and workbook structures must satisfy the documented contracts.
- The self-checks and the output routing have not been run for this snapshot.

The [integration guide](INTEGRATION.md) describes the component placement, the
144-control Word template contract, the export procedure, and focused checks for
these boundaries. Full Office integration and release acceptance remain
unverified.
