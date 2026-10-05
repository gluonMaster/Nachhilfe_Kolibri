# Integration guide

This code belongs to an existing Excel workbook system. Importing it into an empty
workbook will not recreate the worksheets, controls, templates, or data contracts.
Prepare a compatible test workbook with synthetic records before running macros.

## Component placement and encoding

The textual source files are UTF-8. The classic VBA editor and its file importer
can use the VBA project's Windows code page instead. Importing UTF-8 as another
encoding can corrupt comments, worksheet names, and user-facing strings.

For standard and class modules, use a UTF-8-aware editor to read the source. One
option is to create a component of the correct type and name in the VBA editor,
then paste its code body. Omit export metadata such as `Attribute` lines and the
class export's `VERSION`/property header; these are not executable code. Preserve
the component name and any relevant component properties. Check non-ASCII text
after insertion and after exporting it again: clipboard transfer alone is not
proof that the destination code page preserved every character.

Alternatively, convert a separate import copy to the VBA project's code page
using strict, lossless conversion, then use **File > Import File**. Keep the
repository's UTF-8 sources unchanged. If any character is not representable, stop
and resolve the encoding issue instead of silently replacing it with `?`.

| Source | Where it belongs |
| --- | --- |
| `vba/modules/*.bas` | Standard modules, retaining the declared module names |
| `vba/classes/*.cls` | Class modules, retaining class names and export properties |
| `vba/forms/frmProgress.frm` + `.frx` | One UserForm; keep both files together |
| `vba/worksheets/Kinder.vba` | Code pane of the worksheet named `Kinder` |
| `vba/worksheets/Kinder_Blanks.vba` | Code pane of the worksheet named `Kinder_Blanks` |

The worksheet `.vba` files contain code bodies without export headers. Do not
import them as standard modules: their `Worksheet_Change` handlers belong to the
actual worksheet components. If handlers already exist, compare and reconcile
them rather than appending duplicate procedures.

For `frmProgress`, the `.frm` designer information and matching `.frx` are needed
to reconstruct the controls. Do not strip its designer header or treat its text
as a substitute for the entire form. If converting the textual `.frm` before
import, preserve the component/resource names and leave the `.frx` binary intact.
Where a compatible form already exists in the test workbook, compare its code
and controls before replacing it. The retained form's binary design has not been
verified against the current reference workbook.

The optional `ModuleManager` and `VBA_Export_All` utilities access the VBA project
programmatically and require Excel's **Trust access to the VBA project object
model** setting. This is separate from enabling macros. Manual installation
through the VBA editor does not require these utilities. `ModuleManager` is not
an encoding converter; prepare correctly encoded import copies first.

## Workbook contracts

The principal record sheets are `Kinder` and `Kinder_Blanks`. Record processing
starts at row 5. The form-generation path reads `Kinder_Blanks`; it uses that
sheet's `T2` as the reference month/year. An active record has a nonempty surname
in column `C` whose font does not use gray `ColorIndex = 15`.

| Column | Meaning used by the form workflow |
| --- | --- |
| `B` | Family identifier |
| `C`, `D` | Surname and given name |
| `E` | Subject |
| `G`, `H` | Approval-period start and end |
| `I` | Teacher |
| `L` | Date of birth |
| `O` | Benefit reference number, stored as text |
| `S`, `T` | Address fields |

Preserve leading zeroes and punctuation in column `O`. Dates in `G/H` and `T2`
are dates displayed as `dd.mm.yyyy`; formatted identifiers must not be converted
into dates or floating-point representations.

Administrative synchronization also depends on the `Kartei` layout and the
workbook settings in `Kind_Config`. Monthly value transfer has a separate
contract in `ubertNH_Config`. Inspect these modules before adapting workbook
names or column positions. Other workflows require additional sheets such as
`Archiv`, `ErrorLog`, and monthly worksheets; the table above is not a complete
workbook schema.

The separate full AdminSync operation updates existing `Kinder` records from the
administrative workbook, removes records marked `KN`, and synchronizes prepared
form records. An experimental extension that adds missing records directly from
the administrative workbook is not included in this snapshot.

## Document templates

Templates are intentionally absent from the repository. The code expects a
compatible Word template and a legacy Excel template supplied locally.

The Word contract in `NHblank_WordTemplate` requires **144 uniquely tagged content
controls**: eight header/signature controls and 17 attendance rows of eight
controls each. The filled document is expected to remain one page.

The eight non-row tags are:

```text
BG_NR
SCHUELER_NAME
GEBURTSDATUM
UNTERSCHRIFT_ERZIEHUNGSBERECHTIGTE
ANBIETER_NAME_FIRMA
ANBIETER_STEMPEL_UNTERSCHRIFT_DATUM
LEHRKRAFT_NAME
LEHRKRAFT_UNTERSCHRIFT
```

Attendance tags run from `R01_` through `R17_`, each with these suffixes:

```text
DATUM UHRZEIT DAUER GR_EZ FACH ORT STATUS ABWESENHEITSGRUND
```

Each required tag must occur exactly once. The older 127-control contract is
obsolete; it omits the 17 `GR_EZ` controls. The generator applies 10-point,
single-spaced formatting to the reference number, student name, birth date, and
teacher name fields.

The legacy Excel template must contain a `Muster` worksheet. The generator writes
the student name to `B1`, subject to `E1`, approval-period text to `B2`, reference
month to `C4`, and year to `E4`. Recreating these cells alone does not reproduce
the required printable layout.

`NHblank_DataProcessor` currently sets
`GENERATE_WORD_FOR_NON_JOBCENTER = False`: Jobcenter records produce DOCX;
non-Jobcenter/Sozialamt records produce the legacy XLSX form. The record processor
creates Word lazily for DOCX output. Template selection and diagnostic entry
points still need their own review before assuming that a complete workflow can
run on a computer without Word.

## Entry points and focused checks

The wrappers in `NHblank_Menu` expose form generation, record synchronization,
selection transfer, reference-date deactivation, and full administrative sync.
Assign only the required wrappers to controls in the test workbook. Reporting
also depends on `frmProgress`.

After integrating components, run **Debug > Compile VBAProject** in the test
workbook before invoking workflows. Then use the existing checks selectively:

| Function | Purpose and side effects |
| --- | --- |
| `NHblank_BgNummer.NHblank_BgNormalizationSelfTest()` | In-memory reference-number cases; returns `"OK"` on success |
| `NHblank_DataProcessor.NHblank_GenerationSelfTest(wordTemplatePath, legacyTemplatePath, outputFolder)` | Produces one synthetic DOCX and one synthetic XLSX; requires Excel, Word, both templates, and a disposable output folder; returns `"OK"` on success |

The generation self-check can overwrite its fixed test filenames. Use a fresh
disposable output folder. It exercises the two output helpers directly; it does
not prove that the record-processing route correctly selects XLSX-only output.

For the source change that is ahead of the reference workbook, check one
synthetic Jobcenter record and one synthetic Sozialamt record through the actual
generation entry point. Check the resulting output types, then inspect one
representative record edit and report generation to confirm worksheet events and
the progress form are connected correctly. These are suggested integration
checks, not previously completed test results.
