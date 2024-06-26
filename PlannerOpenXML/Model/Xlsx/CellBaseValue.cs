using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PlannerOpenXML.Model.Xlsx;

public abstract class CellBaseValue
{
    #region properties
    public abstract Type Type { get; }
    public abstract object? Value { get; }
    #endregion properties

    #region methods
    internal void Update(XlsxFile xlsxFile, WorkbookPart workbookPart, WorksheetPart worksheetPart, CellReference cellReference)
    {
        var cell = InsertCellInWorksheet(cellReference, worksheetPart);
        Update(xlsxFile, workbookPart, worksheetPart, cell);
    }

    internal abstract void Update(XlsxFile xlsxFile, WorkbookPart workbookPart, WorksheetPart worksheetPart, Cell cell);
    #endregion methods

    #region private methods
    protected internal static Cell InsertCellInWorksheet(CellReference cellReference, WorksheetPart worksheetPart)
    {
        var worksheet = worksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>();

        var row = Sheet.GetOrAddRow(sheetData, cellReference.Row);

        // If there is not a cell with the specified column name, insert one.

        var reference = cellReference.ToString();
        return Sheet.GetOrAddCell(row, reference, cellReference.Column);
    }
    #endregion private methods
}
