using OutSystems.ExternalLibraries.SDK;

namespace BudgetExport;

/// <summary>
/// One row of the Budget sheet. Mirrors the app structure BudgetExportRow.
/// FinalSA, FinalMeritElement and FinalTSI are fractions (0.027 = 2.7%).
/// FinalCPI and CPIHalf are percentage points (3.0 = 3%).
/// </summary>
[OSStructure(Description = "One row of the Budget sheet (one per salary plan).")]
public struct BudgetSheetRow
{
    [OSStructureField(Length = 50, Description = "Salary Admin Plan code")]
    public string SalaryPlan;

    [OSStructureField(Length = 100, Description = "Board (country) name")]
    public string BoardName;

    [OSStructureField(Length = 10, Description = "Y or empty")]
    public string IrregularMarket;

    [OSStructureField(Length = 10, Description = "Y or empty")]
    public string FCS;

    [OSStructureField(Length = 200, Description = "IMF or the alternative source label")]
    public string CPISource;

    [OSStructureField(DataType = OSDataType.Date, Description = "Latest CPI month. NullDate = empty cell")]
    public DateTime LatestCPIMonth;

    [OSStructureField(Length = 37, Decimals = 8, Description = "Final CPI, percentage points")]
    public decimal FinalCPI;

    [OSStructureField(Length = 37, Decimals = 8, Description = "CPI 50%, percentage points")]
    public decimal CPIHalf;

    [OSStructureField(Length = 37, Decimals = 8, Description = "Final SA, fraction")]
    public decimal FinalSA;

    [OSStructureField(Length = 37, Decimals = 8, Description = "Final Merit Element, fraction")]
    public decimal FinalMeritElement;

    [OSStructureField(Length = 37, Decimals = 8, Description = "Final TSI, fraction")]
    public decimal FinalTSI;

    [OSStructureField(Length = 50, Description = "Approach (model type label)")]
    public string Approach;

    [OSStructureField(Length = 2000, Description = "Measure applied")]
    public string MeasureApplied;

    [OSStructureField(Description = "Total staff count WBG")]
    public int StaffCount;

    [OSStructureField(Length = 37, Decimals = 8, Description = "Staff salary WBG in USD")]
    public decimal StaffSalaryUSD;

    [OSStructureField(Length = 37, Decimals = 8, Description = "SA estimate WBG in USD")]
    public decimal SAEstimateUSD;
}
