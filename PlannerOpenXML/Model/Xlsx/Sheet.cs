using CommunityToolkit.Mvvm.ComponentModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PlannerOpenXML.Model.Xlsx;

public partial class Sheet : ObservableObject
{
    #region fields
    private readonly XlsxFile m_XlsxFile;
    private readonly WorksheetPart m_WorksheetPart;
    private readonly WorkbookPart m_WorkbookPart;
    private readonly Worksheet m_Worksheet;
    private MergeCells? m_MergeCells;
    #endregion fields

    #region properties
    [ObservableProperty]
    private string m_Name;

    internal DocumentFormat.OpenXml.Spreadsheet.Sheet SheetVar => m_Sheet;
    private readonly DocumentFormat.OpenXml.Spreadsheet.Sheet m_Sheet;

    public IReadOnlyList<RangeReference> MergedCells => m_MergedCells;
    private readonly List<RangeReference> m_MergedCells = [];
    #endregion properties

    #region constructors
    internal Sheet(XlsxFile xlsxFile, WorkbookPart workbookPart, DocumentFormat.OpenXml.Spreadsheet.Sheet sheet)
    {
        m_XlsxFile = xlsxFile;
        m_WorkbookPart = workbookPart;
        m_WorksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        m_Worksheet = m_WorksheetPart.Worksheet;

        m_MergeCells = m_Worksheet.GetFirstChild<MergeCells>();
        if (m_MergeCells is not null)
        {
            foreach (var cell in m_MergeCells.Elements<MergeCell>())
            {
                if (cell?.Reference?.HasValue == true && cell.Reference.Value is not null)
                    m_MergedCells.Add(new(cell.Reference.Value));
            }
        }

        m_Sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet
        {
            Name = sheet.Name,
            Id = m_WorkbookPart.GetIdOfPart(m_WorksheetPart),
            SheetId = (uint)m_WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().ChildElements.Count + 1
        };
        m_Name = m_Sheet?.Name?.Value ?? string.Empty;
    }

    internal Sheet(XlsxFile xlsxFile, WorkbookPart sourceWorkbookPart, DocumentFormat.OpenXml.Spreadsheet.Sheet sourceSheet, WorkbookPart destinationWorkbookPart)
    {
        m_XlsxFile = xlsxFile;
        m_WorkbookPart = destinationWorkbookPart;

        var sourceWorksheetPart = (WorksheetPart)sourceWorkbookPart.GetPartById(sourceSheet.Id);

        m_WorksheetPart = m_WorkbookPart.AddNewPart<WorksheetPart>();

        m_Worksheet = new Worksheet(new SheetData());
        m_WorksheetPart.Worksheet = m_Worksheet;
        m_Worksheet.Load(sourceWorksheetPart);
        m_Worksheet.GetFirstChild<SheetViews>()?.Remove();

        foreach (var part in sourceWorksheetPart.Parts)
        {
            if (part.OpenXmlPart is SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart)
            {
                m_WorksheetPart.AddPart(spreadsheetPrinterSettingsPart, part.RelationshipId);
            }
        }

        m_Sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet
        {
            Name = sourceSheet.Name,
            Id = m_WorkbookPart.GetIdOfPart(m_WorksheetPart),
            SheetId = (uint)m_WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().ChildElements.Count + 1
        };
        m_Name = m_Sheet?.Name?.Value ?? string.Empty;

        m_Worksheet.Save();
    }
    #endregion constructors

    #region methods
    public void Merge(RangeReference range)
    {
        m_MergedCells.Add(range);
        if (m_MergeCells is null)
        {
            m_MergeCells = new();
            m_Worksheet.InsertAfter(m_MergeCells, m_WorksheetPart.Worksheet.Elements<SheetData>().First());
        }
        m_MergeCells.Append(new MergeCell { Reference = range.ToString() });
    }

    public void SetColumnWidth(uint columnIndex, double width)
    {
        var column = GetOrAddColumn(columnIndex);
        column.Width = width;
        column.CustomWidth = true;
    }

    public void SetRowHeight(uint rowIndex, double height)
    {
        var sheetData = m_Worksheet.GetFirstChild<SheetData>();
        var row = GetOrAddRow(sheetData, rowIndex);

        row.Height = height;
        row.CustomHeight = true;
    }

    public Sheet Clone()
    {
        var result = new Sheet(m_XlsxFile, m_WorkbookPart, m_Sheet, m_WorkbookPart);
        return result;
    }

    public bool TryGetDouble(CellReference cellReference, out double value)
    {
        var cell = GetSpreadsheetCell(m_Worksheet, cellReference);
        if (cell is null || cell.CellValue is null)
        {
            value = 0;
            return false;
        }

        return cell.CellValue.TryGetDouble(out value);
    }

    public void SetValue(CellReference cellReference, CellBaseValue value)
    {
        value.Update(m_XlsxFile, m_WorkbookPart, m_WorksheetPart, cellReference);
    }

    public uint? GetCellStyleIndex(CellReference cellReference)
    {
        var cell = GetSpreadsheetCell(m_Worksheet, cellReference);
        return (cell?.StyleIndex) ?? null;
    }

    public void Save()
    {
        m_WorksheetPart.Worksheet.Save();
        m_WorkbookPart.SharedStringTablePart?.SharedStringTable.Save();
    }

    internal static Cell? GetSpreadsheetCell(Worksheet worksheet, CellReference cellReference)
    {
        var row = worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == cellReference.Row);
        if (row == null)
        {
            // A cell does not exist at the specified row.
            return null;
        }

        var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference.Value, cellReference.ToString(), StringComparison.OrdinalIgnoreCase));
        // A cell does not exist at the specified column, in the specified row.
        return cell;
    }

    internal static Row GetOrAddRow(SheetData sheetData, uint rowIndex)
    {
        // If the worksheet does not contain a row with the specified row index, insert one.
        if (sheetData.ChildElements.Count > rowIndex - 1)
        {
            var testRow = sheetData.ChildElements[(int)rowIndex - 1] as Row;
            if (testRow?.RowIndex == rowIndex)
            {
                return testRow;
            }
        }

        var row = sheetData.Elements<Row>().Reverse().SingleOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.AppendChild(row);
        }
        return row;
    }

    internal Column GetOrAddColumn(uint columnIndex)
    {
        var columns = m_Worksheet.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            var data = m_Worksheet.GetFirstChild<SheetData>();
            m_Worksheet.InsertBefore(columns, data);
        }

        var result = columns.OfType<Column>().FirstOrDefault(x => x.Min == columnIndex && x.Max == columnIndex);
        if (result == null)
        {
            result = new()
            {
                Min = columnIndex,
                Max = columnIndex,
            }            ;
            columns.Append(result);
        }
        return result;
    }

    /// <summary>
    /// Gets or creates a new cell. Bug identified!
    /// TODO: when adding a AA4 cell before Z4 cell the order gets wrong. Need to find a solution.
    /// </summary>
    /// <param name="row"></param>
    /// <param name="cellReference"></param>
    /// <param name="colIndex"></param>
    /// <returns></returns>
    internal static Cell GetOrAddCell(Row row, string cellReference, uint colIndex)
    {
        if (row.ChildElements.Count > colIndex - 1)
        {
            var testCell = row.ChildElements[(int)colIndex - 1] as Cell;
            if (testCell?.CellReference.Value == cellReference)
                return testCell;
        }

        var cellWithColumn = row.Elements<Cell>().Reverse().SingleOrDefault(c => c.CellReference.Value == cellReference);
        if (cellWithColumn != null)
            return cellWithColumn;

        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
        Cell? refCell = null;
        foreach (var cell in row.Elements<Cell>())
        {
            if (cell.CellReference.Value.Length == cellReference.Length)
            {
                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }
        }

        var newCell = new Cell { CellReference = cellReference };
        row.InsertBefore(newCell, refCell);
        return newCell;
    }

    public static void PreFillToRow(WorksheetPart worksheetPart, RangeReference rangeReference)
    {
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        var row = Sheet.GetOrAddRow(sheetData, rangeReference.To.Row);

        if (row.ChildElements.Count == 0)
        {
            for (var i = rangeReference.From.Column; i <= rangeReference.To.Column; i++)
            {
                var cellReference = new CellReference(i, rangeReference.To.Row);
                row.AppendChild(new Cell { CellReference = cellReference.ToString() });
            }
        }
    }
    #endregion methods
}
