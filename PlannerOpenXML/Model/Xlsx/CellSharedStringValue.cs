using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace PlannerOpenXML.Model.Xlsx;

public class CellSharedStringValue(string value, uint? styleIndex = default) : CellBaseValue
{
    #region fields
    private readonly uint? m_StyleIndex = styleIndex;
    #endregion fields

    #region properties
    public override Type Type => typeof(string);

    public override object Value => m_Value;

    public string StringValue => m_Value;
    private readonly string m_Value = value;

    #endregion properties

    #region methods
    internal override void Update(XlsxFile xlsxFile, WorkbookPart workbookPart, WorksheetPart worksheetPart, Cell cell)
    {
        var index = xlsxFile.SharedStringCache.GetIndex(m_Value ?? string.Empty);

        cell.CellValue = new CellValue(index.ToString());
        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        cell.StyleIndex = m_StyleIndex;
    }
    #endregion methods
}
