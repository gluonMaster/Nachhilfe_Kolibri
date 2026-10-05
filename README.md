# Nachhilfe Kolibri

Excel VBA source code for tutoring administration: validating teaching hours,
maintaining student records, preparing monthly worksheets, producing reports,
and generating attendance forms for Jobcenter and Sozialamt workflows.

This repository is a **source snapshot for an existing workbook system**.
Production workbooks, personal records, document templates, and generated reports
are not included. A clean clone is useful for reviewing the implementation; it
does not provide a ready-to-run Excel application.

## What the code covers

- Teaching-hour checks, including weeks that span two months.
- Monthly worksheet preparation and reports covering multiple approval periods.
- Separate working records for form preparation, with synchronization between
  `Kinder` and `Kinder_Blanks`, duplicate handling, and identifier formatting.
- Synchronization with an administrative workbook and transfer of monthly values.
- Editable Word attendance forms for Jobcenter and legacy Excel forms for
  Sozialamt, with validation of benefit reference numbers.
- Worksheet event handlers and the progress form required by the reporting code.

The workflows reflect a particular workbook layout and German administrative
processes. Sheet names, column mappings, and some workbook names remain
integration-specific; they must be reviewed when adapting the code.

## Repository layout

| Path | Contents |
| --- | --- |
| [vba/modules/](vba/modules/) | 52 standard VBA modules |
| [vba/classes/](vba/classes/) | Four VBA class modules |
| [vba/forms/](vba/forms/) | `frmProgress.frm` and its companion `.frx` resource |
| [vba/worksheets/](vba/worksheets/) | Event code for `Kinder` and `Kinder_Blanks` |
| [docs/INTEGRATION.md](docs/INTEGRATION.md) | Component installation, workbook contracts, and focused checks |
| [docs/SOURCE_STATUS.md](docs/SOURCE_STATUS.md) | Source selection and verification limits |

## Requirements and integration

The target environment is desktop Microsoft Excel on Windows with VBA enabled.
DOCX generation additionally uses desktop Microsoft Word. COM dependencies use
late binding, including `Scripting.Dictionary`, `Scripting.FileSystemObject`,
`VBScript.RegExp`, `ADODB.Stream`, and `WScript.Shell`.

Integration requires a compatible workbook and suitable local templates. Review
the [integration guide](docs/INTEGRATION.md) before importing anything. In
particular, the UTF-8 source files need encoding-aware handling in the VBA editor,
and worksheet event code must be placed in the corresponding worksheet modules.

Use a disposable workbook with synthetic records for adaptation and validation.
The repository does not include a complete demonstration workbook or an automated
Office test environment.

## Current status

The source set has been compared offline with the reference workbook. It includes
worksheet code and a progress-form dependency that were missing from the exported
module folder alone.

One source change is ahead of the reference workbook: non-Jobcenter records now
produce only the legacy XLSX form by default, and the record-processing routine
starts Word only when it needs DOCX output. This change is implemented in the
source and documented, but has not been verified by running the updated workbook.
It does not establish that every application entry point can run without Word.

Existing self-checks cover reference-number normalization and synthetic form
generation. They have not been executed as part of this publication preparation.
See [source status](docs/SOURCE_STATUS.md) for the remaining limits, including the
unverified binary design of the progress form.

No license has been assigned to this repository.
