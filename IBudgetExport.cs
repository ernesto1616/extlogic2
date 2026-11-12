using OutSystems.ExternalLibraries.SDK;

namespace BudgetExport;

/// <summary>
/// External Logic library exposed to ODC. One public action that builds the
/// annual Budget workbook (sheet "FYxx Budget WBG") and returns it as XLSX binary.
/// </summary>
[OSInterface(
    Name = "BudgetExport",
    Description = "Generates the annual WBG Compensation Budget workbook (sheet 'FYxx Budget WBG').")]
public interface IBudgetExport
{
    [OSAction(
        Description = "Builds the Budget workbook from one row per salary plan and returns the XLSX file.",
        ReturnName = "ExcelFile",
        ReturnDescription = "XLSX file content.")]
    byte[] GenerateBudgetWorkbook(
        [OSParameter(Description = "Fiscal year label, e.g. FY27. Used in the sheet name, title and group headers.")]
        string fiscalYearLabel,
        [OSParameter(DataType = OSDataType.Date, Description = "Staff snapshot date shown as 'as of' in the Staff Count header. NullDate = not shown.")]
        DateTime staffAsOfDate,
        [OSParameter(Description = "Budget rows, one per salary plan, already sorted.")]
        List<BudgetSheetRow> rows);
}
