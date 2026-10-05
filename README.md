# Tutoring administration with Excel VBA

Excel VBA modules developed for tutoring administration at Kolibri. The code supports the practical workflow around student records, monthly teaching hours, staff workbook copies, and reports used for billing.

This repository contains exported VBA source modules and a progress form. It is a source-code reference for a workbook-based application: the operational workbook, its templates, and student records are not included.

## Main workflows

- Import and reconcile student identifiers and address information from an external workbook.
- Populate monthly worksheets and validate recorded teaching hours.
- Detect duplicate records and maintain an archive.
- Create workbook copies for staff and synchronize their updates.
- Generate individual reports from workbook templates.

## Code map

| Area | Modules |
| --- | --- |
| Student imports and identifiers | [ImportKinder.bas](ImportKinder.bas), [CorrectKinderIdentifiers.bas](CorrectKinderIdentifiers.bas), [CheckExistingNummer.bas](CheckExistingNummer.bas) |
| Address and name processing | [AdressImport.bas](AdressImport.bas), [SplitAddress.bas](SplitAddress.bas), [SurenameNameSplitting.bas](SurenameNameSplitting.bas) |
| Monthly worksheets and hours | [MonatTabelleEinfullung.bas](MonatTabelleEinfullung.bas), [ValidateStudyHours.bas](ValidateStudyHours.bas), [KorrekturStudyload.bas](KorrekturStudyload.bas) |
| Staff copies and synchronization | [ExemplareMachen.bas](ExemplareMachen.bas), [Synchronisation.bas](Synchronisation.bas) |
| Reports and progress display | [Berichten.bas](Berichten.bas), [frmProgress.frm](frmProgress.frm), [frmProgress.frx](frmProgress.frx) |
| Data review and archiving | [Analyzierung.bas](Analyzierung.bas), [RemoveDuplicates.bas](RemoveDuplicates.bas), [Archivierung.bas](Archivierung.bas) |
| Workbook forms and maintenance | [Form.bas](Form.bas), [LoadTableClear.bas](LoadTableClear.bas), [VBA_Export_All.bas](VBA_Export_All.bas) |

## Requirements and integration

The intended environment is desktop Microsoft Excel on Windows with VBA support. The code uses Excel's object model, file dialogs, and Windows scripting objects; it is not an Excel Online add-in or a standalone executable.

To inspect or adapt the modules:

1. Work in a separate macro-enabled test workbook (`.xlsm`) with synthetic records.
2. Open the VBA editor with **Alt+F11** and import the required `.bas` modules using **File > Import File**.
3. If a workflow uses the progress form, keep `frmProgress.frm` and `frmProgress.frx` together and import the `.frm` file.
4. Read the selected procedure and prepare its required sheets, columns, and templates before running it.
5. Compile the VBA project and exercise the selected workflow on the test workbook before adapting it to another workbook.

The original workbook layout is part of the application contract. Examples include sheets named `Kinder`, `Kinder_pre`, `Archiv`, `Form`, `Shablon`, and `Shablon2`, and an external `Kartei` sheet. Several routines act on `ActiveSheet` or use fixed row and column positions. Creating empty sheets with these names alone is not sufficient to reproduce the complete application.

`VBA_Export_All.bas` exports source components and uses Excel's VBA project object model. It is a maintenance utility, not a required step for normal data processing.

## Scope and limitations

- The public snapshot is a collection of workflow-specific modules, not a general-purpose tutoring management package.
- A standalone demo workbook and automated test harness are not included.
- Some procedures update, clear, or archive worksheet data. Review their target sheets and use a disposable test copy when adapting them.
- Personal records, operational workbooks, generated reports, and billing documents belong outside the source repository.
- Existing module names are retained to preserve the mapping to the exported VBA project.
