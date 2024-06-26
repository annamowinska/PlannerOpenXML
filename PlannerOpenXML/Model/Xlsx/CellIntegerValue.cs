using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace PlannerOpenXML.Model.Xlsx;

internal class CellIntegerValue : CellBaseValue
{
    #region fields
    private readonly uint? m_StyleIndex;
    #endregion fields

    #region properties
    public override Type Type => typeof(int);

    public override object Value => m_Value;

    public int DateTimeValue => m_Value;
    private readonly int m_Value;
    #endregion properties

    #region constructors
    public CellIntegerValue(int value, uint? styleIndex = default)
    {
        m_Value = value;
        m_StyleIndex = styleIndex;
    }
    #endregion constructors

    #region methods
    internal override void Update(XlsxFile xlsxFile, WorkbookPart workbookPart, WorksheetPart worksheetPart, Cell cell)
    {
        cell.CellValue = new CellValue(m_Value.ToString());
        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        cell.StyleIndex = m_StyleIndex;
    }
    #endregion methods
}
