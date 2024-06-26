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
    public uint? Day { get; private set; }
    public uint? Saturday { get; private set; }
    public uint? Sunday { get; private set; }
    public double Column1Width { get; private set; } = 8;
    public double Column2Width { get; private set; } = 10;
    public double Column3Width { get; private set; } = 8;
    public double Row1Height { get; private set; } = 70;
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
        Day = template.GetCellStyleIndex(new("B3"));
        Saturday = template.GetCellStyleIndex(new("B4"));
        Sunday = template.GetCellStyleIndex(new("B5"));

        if (template.TryGetDouble(new("E1"), out var column1Width))
            Column1Width = column1Width;
        if (template.TryGetDouble(new("E2"), out var column2Width))
            Column2Width = column2Width;
        if (template.TryGetDouble(new("E3"), out var column3Width))
            Column3Width = column3Width;
        if (template.TryGetDouble(new("E4"), out var row1Height))
            Row1Height = row1Height;

        m_Default = GetTableStyle(template, 8);
        m_Holiday1 = GetTableStyle(template, 12);
        m_Holiday2 = GetTableStyle(template, 16);
        m_Holiday12 = GetTableStyle(template, 20);
        m_Milestone = GetTableStyle(template, 24);

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
