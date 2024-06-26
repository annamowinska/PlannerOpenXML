using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace PlannerOpenXML.Model.Xlsx;

public class CellEmptyValue : CellBaseValue
{
    #region fields
    private readonly uint? m_StyleIndex;
    #endregion fields

    #region properties
    public override Type Type => typeof(object);

    public override object? Value => null;
    #endregion properties

    #region constructors
    public CellEmptyValue()
    {
    }

    public CellEmptyValue(uint? styleIndex = default)
    {
        m_StyleIndex = styleIndex;
    }
    #endregion constructors

    #region methods
    internal override void Update(XlsxFile xlsxFile, WorkbookPart workbookPart, WorksheetPart worksheetPart, Cell cell)
    {
        cell.CellValue?.Remove();
        cell.DataType = null;
        cell.StyleIndex = m_StyleIndex;
    }
    #endregion methods
}
