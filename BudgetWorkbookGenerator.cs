using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace BudgetExport;

/// <summary>
/// Implementation of IBudgetExport. Builds the XLSX with the .NET base library only
/// (System.IO.Compression + hand-written SpreadsheetML), no third-party packages.
///
/// Layout (values only, no formulas):
///   row 2  title
///   row 4  group headers   row 5  sub-group headers   row 6  column headers
///   row 7+ data, then one blank row and the Total row
/// Columns: A empty margin, B..Q = template columns B D G H J K L M N O P Q R S X AC.
/// </summary>
public class BudgetWorkbookGenerator : IBudgetExport
{
    // ---- layout -------------------------------------------------------------
    private const int TitleRow = 2;
    private const int GroupRow = 4;
    private const int SubGroupRow = 5;
    private const int HeaderRow = 6;
    private const int FirstDataRow = 7;

    private const int ColPlan = 2;       // B  Salary Plan
    private const int ColBoard = 3;      // D  Board Name
    private const int ColIrregular = 4;  // G  Irregular Market
    private const int ColFcs = 5;        // H  FCS?
    private const int ColCpiSource = 6;  // J  CPI Source
    private const int ColCpiMonth = 7;   // K  Latest CPI Month
    private const int ColCpi = 8;        // L  Final CPI
    private const int ColCpiHalf = 9;    // M  CPI (50%)
    private const int ColSa = 10;        // N  Final SA
    private const int ColMe = 11;        // O  Final Merit Element
    private const int ColTsi = 12;       // P  Final TSI
    private const int ColApproach = 13;  // Q  Approach
    private const int ColMeasure = 14;   // R  Measure Applied
    private const int ColStaff = 15;     // S  Staff Count WBG
    private const int ColSalary = 16;    // X  Staff Salary WBG (USD)
    private const int ColSaEst = 17;     // AC SA Estimate WBG (USD)
    private const int LastCol = ColSaEst;

    // ---- style indexes (cellXfs order in StylesXml) -------------------------
    private const int SDefault = 0;
    private const int STitle = 1;
    private const int SGroup = 2;
    private const int SSubGroup = 3;
    private const int SHeader = 4;
    private const int STextCenter = 5;
    private const int STextLeft = 6;
    private const int STextWrap = 7;
    private const int SMonth = 8;
    private const int SPoints = 9;
    private const int SPct1 = 10;
    private const int SPct2 = 11;
    private const int SCount = 12;
    private const int SUsd = 13;
    private const int STotalLabel = 14;
    private const int STotalCount = 15;
    private const int STotalUsd = 16;

    public byte[] GenerateBudgetWorkbook(string fiscalYearLabel, DateTime staffAsOfDate, List<BudgetSheetRow> rows)
    {
        var fy = string.IsNullOrWhiteSpace(fiscalYearLabel) ? "FY" : fiscalYearLabel.Trim();
        rows ??= new List<BudgetSheetRow>();

        var sheet = new SheetBuilder();

        // Title
        sheet.Text(TitleRow, ColPlan, $"{fy} World Bank Group Compensation Budget", STitle);

        // Group headers (row 4) and sub-group headers (row 5)
        var staffLabel = IsNullDate(staffAsOfDate)
            ? "Staff Count"
            : $"Staff Count (as of {staffAsOfDate.ToString("MM.dd.yyyy", CultureInfo.InvariantCulture)})";

        sheet.Block(GroupRow, ColPlan, SubGroupRow, ColFcs, "Salary Plan Data", SGroup);
        sheet.Block(GroupRow, ColCpiSource, GroupRow, ColMeasure, $"{fy} Review Data", SGroup);
        sheet.Block(GroupRow, ColStaff, GroupRow, ColSalary, "Staff Data", SGroup);
        sheet.Block(GroupRow, ColSaEst, GroupRow, ColSaEst, $"{fy} Budget Outcome SA", SGroup);

        sheet.Block(SubGroupRow, ColCpiSource, SubGroupRow, ColCpiHalf, "CPI", SSubGroup);
        sheet.Block(SubGroupRow, ColSa, SubGroupRow, ColMeasure, "SA, ME and TSI", SSubGroup);
        sheet.Block(SubGroupRow, ColStaff, SubGroupRow, ColStaff, staffLabel, SSubGroup);
        sheet.Block(SubGroupRow, ColSalary, SubGroupRow, ColSalary, "Staff Salary (USD)", SSubGroup);
        sheet.Block(SubGroupRow, ColSaEst, SubGroupRow, ColSaEst, "SA Estimate (USD)", SSubGroup);

        // Column headers (row 6)
        sheet.Text(HeaderRow, ColPlan, "Salary Plan", SHeader);
        sheet.Text(HeaderRow, ColBoard, "Board Name", SHeader);
        sheet.Text(HeaderRow, ColIrregular, "Irregular Market", SHeader);
        sheet.Text(HeaderRow, ColFcs, "FCS?", SHeader);
        sheet.Text(HeaderRow, ColCpiSource, "CPI Source", SHeader);
        sheet.Text(HeaderRow, ColCpiMonth, "Latest CPI\nMonth", SHeader);
        sheet.Text(HeaderRow, ColCpi, "Final CPI", SHeader);
        sheet.Text(HeaderRow, ColCpiHalf, "CPI (50%)", SHeader);
        sheet.Text(HeaderRow, ColSa, "Final SA", SHeader);
        sheet.Text(HeaderRow, ColMe, "Final Merit Element", SHeader);
        sheet.Text(HeaderRow, ColTsi, "Final TSI", SHeader);
        sheet.Text(HeaderRow, ColApproach, "Approach", SHeader);
        sheet.Text(HeaderRow, ColMeasure, "Measure Applied", SHeader);
        sheet.Text(HeaderRow, ColStaff, "WBG", SHeader);
        sheet.Text(HeaderRow, ColSalary, "WBG", SHeader);
        sheet.Text(HeaderRow, ColSaEst, "WBG", SHeader);
        sheet.RowHeight(HeaderRow, 43.15);

        // Data rows
        var r = FirstDataRow;
        foreach (var d in rows)
        {
            sheet.Text(r, ColPlan, d.SalaryPlan, STextCenter);
            sheet.Text(r, ColBoard, d.BoardName, STextLeft);
            sheet.Text(r, ColIrregular, d.IrregularMarket, STextCenter);
            sheet.Text(r, ColFcs, d.FCS, STextCenter);
            sheet.Text(r, ColCpiSource, d.CPISource, STextCenter);
            sheet.Date(r, ColCpiMonth, d.LatestCPIMonth, SMonth);
            sheet.Number(r, ColCpi, d.FinalCPI, SPoints);
            sheet.Number(r, ColCpiHalf, d.CPIHalf, SPoints);
            sheet.Number(r, ColSa, d.FinalSA, SPct1);
            sheet.Number(r, ColMe, d.FinalMeritElement, SPct1);
            sheet.Number(r, ColTsi, d.FinalTSI, SPct2);
            sheet.Text(r, ColApproach, d.Approach, STextCenter);
            sheet.Text(r, ColMeasure, d.MeasureApplied, STextWrap);
            sheet.Number(r, ColStaff, d.StaffCount, SCount);
            sheet.Number(r, ColSalary, d.StaffSalaryUSD, SUsd);
            sheet.Number(r, ColSaEst, d.SAEstimateUSD, SUsd);
            r++;
        }
        var lastDataRow = Math.Max(r - 1, HeaderRow);

        // Total row
        var totalRow = lastDataRow + 2;
        sheet.Text(totalRow, ColMeasure, "Total", STotalLabel);
        sheet.Number(totalRow, ColStaff, rows.Sum(x => (decimal)x.StaffCount), STotalCount);
        sheet.Number(totalRow, ColSalary, rows.Sum(x => x.StaffSalaryUSD), STotalUsd);
        sheet.Number(totalRow, ColSaEst, rows.Sum(x => x.SAEstimateUSD), STotalUsd);

        var filterRef = $"{SheetBuilder.Ref(HeaderRow, ColPlan)}:{SheetBuilder.Ref(lastDataRow, LastCol)}";
        var sheetXml = sheet.ToXml(ColumnWidths(), filterRef, HeaderRow);

        var sheetName = SafeSheetName($"{fy} Budget WBG");
        var absFilterRef = $"'{sheetName.Replace("'", "''")}'!${SheetBuilder.ColLetter(ColPlan)}${HeaderRow}:${SheetBuilder.ColLetter(LastCol)}${lastDataRow}";
        var workbookXml = WorkbookXml
            .Replace("{{SHEET_NAME}}", SheetBuilder.Escape(sheetName))
            .Replace("{{FILTER_REF}}", SheetBuilder.Escape(absFilterRef));

        return Package(new (string, string)[]
        {
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", RootRelsXml),
            ("xl/workbook.xml", workbookXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/styles.xml", StylesXml),
            ("xl/worksheets/sheet1.xml", sheetXml),
        });
    }

    // ---- helpers -----------------------------------------------------------
    private static (int Col, double Width)[] ColumnWidths() => new[]
    {
        (1, 3.57),
        (ColPlan, 8.0), (ColBoard, 23.71), (ColIrregular, 9.0), (ColFcs, 7.0),
        (ColCpiSource, 14.0), (ColCpiMonth, 13.29), (ColCpi, 10.0), (ColCpiHalf, 10.71),
        (ColSa, 9.71), (ColMe, 10.14), (ColTsi, 9.71), (ColApproach, 14.0),
        (ColMeasure, 55.86), (ColStaff, 12.0), (ColSalary, 15.0), (ColSaEst, 15.0),
    };

    // OutSystems NullDate() = 1900-01-01
    internal static bool IsNullDate(DateTime d) => d.Year <= 1900;

    private static string SafeSheetName(string name)
    {
        foreach (var ch in new[] { '[', ']', ':', '*', '?', '/', '\\' })
            name = name.Replace(ch, ' ');
        return name.Length > 31 ? name.Substring(0, 31) : name;
    }

    private static byte[] Package((string Path, string Xml)[] parts)
    {
        var utf8 = new UTF8Encoding(false);
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (path, xml) in parts)
            {
                var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
                using var s = entry.Open();
                var bytes = utf8.GetBytes(xml);
                s.Write(bytes, 0, bytes.Length);
            }
        }
        return ms.ToArray();
    }

    // ---- static package parts ----------------------------------------------
    private const string ContentTypesXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>
""";

    private const string RootRelsXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
""";

    private const string WorkbookRelsXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
""";

    private const string WorkbookXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="{{SHEET_NAME}}" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">{{FILTER_REF}}</definedName></definedNames></workbook>
""";

    // cellXfs order must match the S* constants above.
    private const string StylesXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="4"><numFmt numFmtId="164" formatCode="[$-409]mmmm\-yy;@"/><numFmt numFmtId="165" formatCode="#,##0_);(#,##0)"/><numFmt numFmtId="166" formatCode="\$#,###;\$\-#,###;\$0"/><numFmt numFmtId="167" formatCode="0.0%"/></numFmts><fonts count="6"><font><sz val="9"/><color theme="1"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="14"/><color theme="1"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="8"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font><font><sz val="8"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="8"/><color theme="1"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="9"/><color theme="1"/><name val="Arial"/><family val="2"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF002060"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF002A7E"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="4"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border><border><left/><right/><top style="thin"><color rgb="FF000000"/></top><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="17"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="2" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="167" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="10" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="165" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="166" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf><xf numFmtId="165" fontId="5" fillId="0" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="166" fontId="5" fillId="0" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>
""";
}

/// <summary>Minimal in-memory worksheet writer (inline strings, numbers, dates, merges).</summary>
internal sealed class SheetBuilder
{
    private readonly SortedDictionary<int, SortedDictionary<int, string>> _rows = new();
    private readonly Dictionary<int, double> _heights = new();
    private readonly List<string> _merges = new();

    public void Text(int row, int col, string value, int style)
    {
        var v = Clean(value?.Trim());
        if (string.IsNullOrEmpty(v)) { Blank(row, col, style); return; }
        Put(row, col, $"<c r=\"{Ref(row, col)}\" s=\"{style}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{Escape(v)}</t></is></c>");
    }

    public void Number(int row, int col, decimal value, int style)
        => Put(row, col, $"<c r=\"{Ref(row, col)}\" s=\"{style}\"><v>{value.ToString(CultureInfo.InvariantCulture)}</v></c>");

    public void Date(int row, int col, DateTime value, int style)
    {
        if (BudgetWorkbookGenerator.IsNullDate(value)) { Blank(row, col, style); return; }
        var serial = value.Date.ToOADate().ToString(CultureInfo.InvariantCulture);
        Put(row, col, $"<c r=\"{Ref(row, col)}\" s=\"{style}\"><v>{serial}</v></c>");
    }

    public void Blank(int row, int col, int style)
        => Put(row, col, $"<c r=\"{Ref(row, col)}\" s=\"{style}\"/>");

    /// <summary>Writes text in the top-left cell, styles every cell of the range and merges it.</summary>
    public void Block(int r1, int c1, int r2, int c2, string text, int style)
    {
        for (var r = r1; r <= r2; r++)
            for (var c = c1; c <= c2; c++)
                Blank(r, c, style);
        Text(r1, c1, text, style);
        if (r1 != r2 || c1 != c2) _merges.Add($"{Ref(r1, c1)}:{Ref(r2, c2)}");
    }

    public void RowHeight(int row, double height) => _heights[row] = height;

    public string ToXml((int Col, double Width)[] widths, string autoFilterRef, int freezeRows)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

        var top = Ref(freezeRows + 1, 1);
        sb.Append("<sheetViews><sheetView workbookViewId=\"0\">");
        sb.Append($"<pane ySplit=\"{freezeRows}\" topLeftCell=\"{top}\" activePane=\"bottomLeft\" state=\"frozen\"/>");
        sb.Append($"<selection pane=\"bottomLeft\" activeCell=\"{top}\" sqref=\"{top}\"/>");
        sb.Append("</sheetView></sheetViews>");
        sb.Append("<sheetFormatPr defaultRowHeight=\"12\"/>");

        sb.Append("<cols>");
        foreach (var (col, width) in widths.OrderBy(w => w.Col))
            sb.Append($"<col min=\"{col}\" max=\"{col}\" width=\"{width.ToString(CultureInfo.InvariantCulture)}\" customWidth=\"1\"/>");
        sb.Append("</cols>");

        sb.Append("<sheetData>");
        var allRows = new SortedSet<int>(_rows.Keys);
        allRows.UnionWith(_heights.Keys);
        foreach (var r in allRows)
        {
            sb.Append($"<row r=\"{r}\"");
            if (_heights.TryGetValue(r, out var h))
                sb.Append($" ht=\"{h.ToString(CultureInfo.InvariantCulture)}\" customHeight=\"1\"");
            sb.Append('>');
            if (_rows.TryGetValue(r, out var cells))
                foreach (var cell in cells.Values) sb.Append(cell);
            sb.Append("</row>");
        }
        sb.Append("</sheetData>");

        sb.Append($"<autoFilter ref=\"{autoFilterRef}\"/>");

        if (_merges.Count > 0)
        {
            sb.Append($"<mergeCells count=\"{_merges.Count}\">");
            foreach (var m in _merges) sb.Append($"<mergeCell ref=\"{m}\"/>");
            sb.Append("</mergeCells>");
        }

        sb.Append("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");
        sb.Append("</worksheet>");
        return sb.ToString();
    }

    private void Put(int row, int col, string xml)
    {
        if (!_rows.TryGetValue(row, out var cells))
        {
            cells = new SortedDictionary<int, string>();
            _rows[row] = cells;
        }
        cells[col] = xml;
    }

    public static string Ref(int row, int col) => ColLetter(col) + row.ToString(CultureInfo.InvariantCulture);

    public static string ColLetter(int col)
    {
        var s = "";
        while (col > 0)
        {
            var m = (col - 1) % 26;
            s = (char)('A' + m) + s;
            col = (col - 1) / 26;
        }
        return s;
    }

    public static string Escape(string s) => s
        .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
        .Replace("\"", "&quot;").Replace("'", "&apos;");

    // Removes characters that are not allowed in XML 1.0 (keeps tab, LF, CR).
    private static string Clean(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
            if (ch == '\t' || ch == '\n' || ch == '\r' || ch >= 0x20) sb.Append(ch);
        return sb.ToString();
    }
}
