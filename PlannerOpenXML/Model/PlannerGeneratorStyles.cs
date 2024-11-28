using DocumentFormat.OpenXml.Wordprocessing;
using PlannerOpenXML.Model.Xlsx;

namespace PlannerOpenXML.Model;

internal enum CellStyle
{
    Default = 0,
    Holiday1 = 1,
    Holiday2 = 2,
    Holiday12 = 3,
    Milestone = 4,
}

internal class PlannerGeneratorStyles
{
    #region fields
    private (uint? cell_a1, uint? cell_a2, uint? cell_b1, uint? cell_b2) m_Default;
    private (uint? cell_a1, uint? cell_a2, uint? cell_b1, uint? cell_b2) m_Holiday1;
    private (uint? cell_a1, uint? cell_a2, uint? cell_b1, uint? cell_b2) m_Holiday2;
    private (uint? cell_a1, uint? cell_a2, uint? cell_b1, uint? cell_b2) m_Holiday12;
    private (uint? cell_a1, uint? cell_a2, uint? cell_b1, uint? cell_b2) m_Milestone;
    #endregion fields

    #region properties
    public bool Failed { get; }
    public string FailedReason { get; }

    public uint? Month { get; private set; }
    public uint? Week { get; private set; }
    public uint? MonthDay { get; private set; }
    public uint? WeekDay { get; private set; }
    public uint? MonthDaySaturday { get; private set; }
    public uint? WeekDaySaturday { get; private set; }
    public uint? MonthDaySunday { get; private set; }
    public uint? WeekDaySunday { get; private set; }
    public uint? Year { get; private set; }
    public uint? Header { get; private set; }
    public uint? Footer1 { get; private set; }
    public uint? Footer2 { get; private set; }
    public uint? Footer0 { get; private set; }
    public double Column1Width { get; private set; } = 9;
    public double Column2Width { get; private set; } = 5;
    public double Column3Width { get; private set; } = 18;
    public double Column4Width { get; private set; } = 7.5;
    public double Row1Height { get; private set; } = 350;
    public double Row2Height { get; private set; } = 130;
    public double DayRowHeight { get; private set; } = 40;
    public double Footer0RowHeight { get; private set; } = 20;
    public double Footer1RowHeight { get; private set; } = 100;
    public double Footer2RowHeight { get; private set; } = 200;
    #endregion properties

    #region constructors
    public PlannerGeneratorStyles(XlsxFile xlsxFile)
    {
        (Failed, FailedReason) = ReadStyles(xlsxFile);
    }
    #endregion constructors

    #region methods
    public uint? GetStyleIndex(CellStyle style, uint column, uint row)
    {
        return style switch
        {
            CellStyle.Default => GetStyleIndex(m_Default, column, row),
            CellStyle.Holiday1 => GetStyleIndex(m_Holiday1, column, row),
            CellStyle.Holiday2 => GetStyleIndex(m_Holiday2, column, row),
            CellStyle.Holiday12 => GetStyleIndex(m_Holiday12, column, row),
            CellStyle.Milestone => GetStyleIndex(m_Milestone, column, row),
            _ => throw new NotImplementedException(),
        };
    }
    #endregion methods

    #region private methods
    private static uint? GetStyleIndex((uint? cell_a1, uint? cell_a2, uint? cell_b1, uint? cell_b2) values, uint column, uint row)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, (uint)1, nameof(column));
        ArgumentOutOfRangeException.ThrowIfGreaterThan(column, (uint)2, nameof(column));
        ArgumentOutOfRangeException.ThrowIfLessThan(row, (uint)1, nameof(row));
        ArgumentOutOfRangeException.ThrowIfGreaterThan(row, (uint)2, nameof(row));
        if (column == 1 && row == 1)
            return values.cell_a1;
        if (column == 2 && row == 1)
            return values.cell_b1;
        if (column == 1 && row == 2)
            return values.cell_a2;
        return values.cell_b2;
    }

    private (bool failed, string reason) ReadStyles(XlsxFile xlsxFile)
    {
        var failed = false;
        var reason = string.Empty;

        if (!xlsxFile.Sheets.TryGetByName("Template", out var template))
            return (true, "Template sheet not found");

        Month = template.GetCellStyleIndex(new("B1"));
        Week = template.GetCellStyleIndex(new("B2"));
        MonthDay = template.GetCellStyleIndex(new("B3"));
        WeekDay = template.GetCellStyleIndex(new("B4"));
        MonthDaySaturday = template.GetCellStyleIndex(new("B5"));
        WeekDaySaturday = template.GetCellStyleIndex(new("B6"));
        MonthDaySunday = template.GetCellStyleIndex(new("B7"));
        WeekDaySunday = template.GetCellStyleIndex(new("B8"));
        Year = template.GetCellStyleIndex(new("B9"));
        Header = template.GetCellStyleIndex(new("B10"));
        Footer1 = template.GetCellStyleIndex(new("B11"));
        Footer2 = template.GetCellStyleIndex(new("B12"));
        Footer0 = template.GetCellStyleIndex(new("B13"));

        if (template.TryGetDouble(new("E1"), out var column1Width))
            Column1Width = column1Width;
        if (template.TryGetDouble(new("E2"), out var column2Width))
            Column2Width = column2Width;
        if (template.TryGetDouble(new("E3"), out var column3Width))
            Column3Width = column3Width;
        if (template.TryGetDouble(new("E4"), out var column4Width))
            Column4Width = column4Width;
        if (template.TryGetDouble(new("E5"), out var row1Height))
            Row1Height = row1Height;
        if (template.TryGetDouble(new("E6"), out var row2Height))
            Row2Height = row2Height;
        if (template.TryGetDouble(new("E7"), out var dayRowHeight))
            DayRowHeight = dayRowHeight;
        if (template.TryGetDouble(new("E8"), out var footer0RowHeight))
            Footer0RowHeight = footer0RowHeight;
        if (template.TryGetDouble(new("E9"), out var footer1RowHeight))
            Footer1RowHeight = footer1RowHeight;
        if (template.TryGetDouble(new("E10"), out var footer2RowHeight))
            Footer2RowHeight = footer2RowHeight;

        m_Default = GetTableStyle(template, 16);
        m_Holiday1 = GetTableStyle(template, 20);
        m_Holiday2 = GetTableStyle(template, 24);
        m_Holiday12 = GetTableStyle(template, 28);
        m_Milestone = GetTableStyle(template, 32);

        return (failed, reason);
    }

    private static (uint? cell_a1, uint? cell_a2, uint? cell_b1, uint? cell_b2) GetTableStyle(Sheet sheet, uint startRow)
    {

        var cell_a1 = sheet.GetCellStyleIndex(new(1, startRow));
        var cell_a2 = sheet.GetCellStyleIndex(new(1, startRow + 1));
        var cell_b1 = sheet.GetCellStyleIndex(new(2, startRow));
        var cell_b2 = sheet.GetCellStyleIndex(new(2, startRow + 1));

        return (cell_a1, cell_a2, cell_b1, cell_b2);
    }
    #endregion private methods
}
