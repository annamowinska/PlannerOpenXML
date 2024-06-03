using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

public class SpreadsheetService
{
    #region methods
    public static void AppendCellToWorksheet(WorksheetPart worksheetPart, Cell cell, uint rowIndex, uint columnIndex)
    {
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }
        else
        {
            Cell existingCell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == $"{GetColumnName(columnIndex)}{rowIndex}");

            if (existingCell != null)
            {
                row.RemoveChild(existingCell);
            }
        }
        cell.CellReference = $"{GetColumnName(columnIndex)}{rowIndex}";
        row.Append(cell);
    }

    public static string GetColumnName(uint columnIndex)
    {
        uint dividend = columnIndex;
        string columnName = string.Empty;
        uint modifier;

        while (dividend > 0)
        {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (uint)((dividend - modifier) / 26);
        }

        return columnName;
    }

    public static void SetColumnWidth(WorksheetPart worksheetPart, uint columnIndex, double width)
    {
        var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            worksheetPart.Worksheet.InsertAt(columns, 0);
        }

        var column = new Column()
        {
            Min = columnIndex,
            Max = columnIndex,
            Width = width,
            CustomWidth = true
        };

        columns.Append(column);
    }

    public static void SetRowHeight(WorksheetPart worksheetPart, double height, uint rowIndex)
    {
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);

        if (row == null)
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        row.Height = height;
        row.CustomHeight = true;
    }
    #endregion methods
}