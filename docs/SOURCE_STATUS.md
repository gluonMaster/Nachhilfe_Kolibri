# Source status

This snapshot combines the current exported source modules with the supporting
components needed by the reference workbook. The selection was checked offline;
the preparation process did not run Excel, Word, or VBA macros.

## Included sources

| Component group | Selection and evidence |
| --- | --- |
| 52 standard modules and four classes | Current exported source set, compared with the embedded VBA in the reference workbook |
| `frmProgress.frm` and `.frx` | Retained from the earlier repository because report generation still uses the form; its code matches the reference workbook, but the binary designer resource has not been verified against that workbook |
| `Kinder.vba` and `Kinder_Blanks.vba` | Worksheet event code extracted from the reference workbook; export headers removed so the files can be placed in the corresponding worksheet code panes |

Production workbooks, administrative data, templates, generated output, temporary
Office files, and internal task prompts are excluded. The repository is not a
backup of the operational workbook system.

## Source change ahead of the workbook

`NHblank_DataProcessor` contains an intentional newer behavior that is absent
from the compared workbook copies:

- `GENERATE_WORD_FOR_NON_JOBCENTER` defaults to `False`.
- A non-Jobcenter/Sozialamt record produces the legacy XLSX output only.
- The record-processing routine creates Word only when DOCX output is needed.

The source documentation explicitly describes the new default and identifies
DOCX plus XLSX as the previous behavior. This establishes the intended source
contract. No saved result confirming its import into the reference workbook or
successful runtime verification was found.

The source behavior is retained in this publication snapshot. Its inclusion must
not be read as a claim that the operational workbook already runs this version.
It also does not prove that all diagnostic and template-selection paths work
without Word.

## Deliberately excluded experiment

A separate, uncommitted AdminSync extension can add missing records from the
administrative workbook to `Kinder`, including restoration from `Archiv`.
It is absent from both the selected exported module and the reference workbook.
It has not been merged into this publication snapshot. The included AdminSync
therefore retains the existing-record synchronization workflow.

## Verification limits

Offline source comparison establishes which code was selected, not that every
workflow compiles and runs in a freshly integrated workbook. In particular:

- The binary layout of the retained progress form needs integration verification.
- Worksheet event code must be attached to the correct worksheet components.
- UTF-8 source text must survive the target VBA project's encoding.
- Templates and workbook structures must satisfy the documented contracts.
- The existing self-checks and the newer output-routing behavior have not been
  executed during publication preparation.

The [integration guide](INTEGRATION.md) describes the required component placement,
the current 144-control Word template contract, and focused checks for these
boundaries. Full Office integration and release acceptance remain unverified.
