using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PlannerOpenXML.Services;
using System.Globalization;
using System.Windows;

namespace PlannerOpenXML.Model;

public class PlannerGenerator
{
    #region constructor
    public PlannerGenerator(IApiService apiService, HolidayNameService holidayNameService, PlannerStyleService plannerStyleService)
    {
        m_ApiService = apiService;
        m_HolidayNameService = holidayNameService;
        m_PlannerStyleService = plannerStyleService;
    }
    #endregion constructor

    #region fields
    private readonly IApiService m_ApiService;
    private readonly HolidayNameService m_HolidayNameService;
    private readonly PlannerStyleService m_PlannerStyleService;
    #endregion fields

    #region methods
    /// <summary>
    /// Generates an excel with a calendar for a given time period for german and hungarian holidays.
    /// </summary>
    /// <param name="year">Starting year</param>
    /// <param name="firstMonth">Staring month</param>
    /// <param name="numberOfMonths">Amount of months to generate</param>
    /// <param name="path">Path for destination file</param>
    /// <returns>Nothing. Async task.</returns>
    public async Task GeneratePlanner(int year, int firstMonth, int numberOfMonths, string path)
    {
        try
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Stylesheet workbookstylesheet = m_PlannerStyleService.GenerateStylesheet();

                WorkbookStylesPart stylesheet = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesheet.Stylesheet = workbookstylesheet;

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Planner" };
                sheets.Append(sheet);

                CultureInfo culture = new CultureInfo("de-DE");

                int currentYear = year;
                int currentMonth = firstMonth;

                var germanHolidaysTask = m_ApiService.GetHolidaysAsync(currentYear, "DE");
                var hungarianHolidaysTask = m_ApiService.GetHolidaysAsync(currentYear, "HU");

                var germanHolidays = await germanHolidaysTask;
                var hungarianHolidays = await hungarianHolidaysTask;

                for (int i = 0; i < numberOfMonths; i++)
                {
                    currentYear = year;
                    currentMonth = firstMonth + i;

                    while (currentMonth > 12)
                    {
                        currentYear++;
                        currentMonth -= 12;

                        germanHolidaysTask = m_ApiService.GetHolidaysAsync(currentYear, "DE");
                        hungarianHolidaysTask = m_ApiService.GetHolidaysAsync(currentYear, "HU");

                        germanHolidays = await germanHolidaysTask;
                        hungarianHolidays = await hungarianHolidaysTask;
                    }

                    DateOnly monthDate = new DateOnly(currentYear, currentMonth, 1);
                    string monthName = monthDate.ToString("MMMM yyyy", culture);

                    Cell monthYearCell = new Cell(new CellValue($"{monthName}"))
                    {
                        DataType = CellValues.String,
                        StyleIndex = 1
                    };

                    AppendCellToWorksheet(spreadsheetDocument, worksheetPart, monthYearCell, 1, (uint)(i + 1));

                    int currentRow = 2;

                    DateOnly currentDate = monthDate;
                    while (currentDate.Month == monthDate.Month)
                    {
                        string dayOfWeek = currentDate.ToString("ddd", culture);
                        string cellValue = $"{currentDate.Day} {dayOfWeek}";

                        string germanHolidayName = m_HolidayNameService.GetHolidayName(currentDate, germanHolidays);
                        string hungarianHolidayName = m_HolidayNameService.GetHolidayName(currentDate, hungarianHolidays);

                        DateOnly nextMonth = monthDate.AddMonths(1);
                        bool isLastDayOfMonth = currentDate.AddDays(1).Month != nextMonth.Month;

                        Cell dateCell = new Cell(new CellValue(cellValue))
                        {
                            StyleIndex = 2
                        };

                        if (!string.IsNullOrEmpty(germanHolidayName) && !string.IsNullOrEmpty(hungarianHolidayName))
                            cellValue += $" DE&HU: {germanHolidayName}";
                        else if (!string.IsNullOrEmpty(germanHolidayName))
                            cellValue += $" DE: {germanHolidayName}";
                        else if (!string.IsNullOrEmpty(hungarianHolidayName))
                            cellValue += $" HU: {hungarianHolidayName}";

                        dateCell = new Cell(new CellValue(cellValue))
                        {
                            DataType = CellValues.String,
                            StyleIndex = 2
                        };

                        if (!isLastDayOfMonth)
                        {
                            if (currentDate.DayOfWeek == DayOfWeek.Saturday)
                                dateCell.StyleIndex = 9;
                            else if (currentDate.DayOfWeek == DayOfWeek.Sunday)
                                dateCell.StyleIndex = 10;
                            else if (cellValue.Contains(" DE:"))
                                dateCell.StyleIndex = 11;
                            else if (cellValue.Contains(" HU:"))
                                dateCell.StyleIndex = 12;
                            else if (cellValue.Contains(" DE&HU:"))
                                dateCell.StyleIndex = 13;
                            else
                                dateCell.StyleIndex = 8;
                        }
                        else if (currentDate.DayOfWeek == DayOfWeek.Saturday)
                            dateCell.StyleIndex = 3;
                        else if (currentDate.DayOfWeek == DayOfWeek.Sunday)
                            dateCell.StyleIndex = 4;
                        else if (cellValue.Contains(" DE:"))
                            dateCell.StyleIndex = 5;
                        else if (cellValue.Contains(" HU:"))
                            dateCell.StyleIndex = 6;
                        else if (cellValue.Contains(" DE&HU:"))
                            dateCell.StyleIndex = 7;

                        AppendCellToWorksheet(spreadsheetDocument, worksheetPart, dateCell, (uint)currentRow, (uint)(i + 1));

                        currentDate = currentDate.AddDays(1);
                        currentRow++;
                    }

                    SetColumnWidth(worksheetPart, (uint)(i + 1), 35);
                }

                workbookPart.Workbook.Save();
            }

            MessageBox.Show($"Planner has been generated and saved as: {path}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred while generating the planner: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    #endregion methods

    #region private methods
    private static void AppendCellToWorksheet(SpreadsheetDocument spreadsheetDocument, WorksheetPart worksheetPart, Cell cell, uint rowIndex, uint columnIndex)
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

    private static string GetColumnName(uint columnIndex)
    {
        uint dividend = columnIndex;
        string columnName = String.Empty;
        uint modifier;

        while (dividend > 0)
        {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (uint)((dividend - modifier) / 26);
        }

        return columnName;
    }

    private static void SetColumnWidth(WorksheetPart worksheetPart, uint columnIndex, double width)
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
    #endregion private methods
}
