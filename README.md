# BudgetExport - ODC External Logic

Generates the annual WBG Compensation Budget workbook (sheet `FYxx Budget WBG`).
No third-party packages: the XLSX is written with System.IO.Compression.

Action: `GenerateBudgetWorkbook(fiscalYearLabel, staffAsOfDate, rows) -> ExcelFile (Binary)`

Build: GitHub Actions (`.github/workflows/build.yml`), linux-x64.
Download the artifact `BudgetExport-<run>` and upload it as-is to ODC Portal > External Logic.
