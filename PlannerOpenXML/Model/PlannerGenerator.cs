using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PlannerOpenXML.Services;
using System.Globalization;
using System.Windows;

namespace PlannerOpenXML.Model;

public class PlannerGenerator
{
    #region fields
    private readonly HolidayNameService m_HolidayNameService;
    private readonly PlannerStyleService m_PlannerStyleService;
    #endregion fields

    #region constructors
    public PlannerGenerator(
        IApiService apiService, 
        HolidayNameService holidayNameService,
        PlannerStyleService plannerStyleService)
    {
        m_HolidayNameService = holidayNameService;
        m_PlannerStyleService = plannerStyleService;
    }
    #endregion constructors

    #region methods
    /// <summary>
    /// Generates an excel with a calendar for a given time period for german and hungarian holidays.
    /// </summary>
    /// <param name="year">Starting year</param>
    /// <param name="firstMonth">Staring month</param>
    /// <param name="numberOfMonths">Amount of months to generate</param>
    /// <param name="path">Path for destination file</param>
    /// <returns>Nothing. Async task.</returns>
    public async Task GeneratePlanner(
            DateOnly from, 
            DateOnly to, 
            IEnumerable<Holiday> allHolidays, 
            string path, 
            string firstCountryCode, 
            string secondCountryCode, 
            IEnumerable<Milestone> milestones)
    {
        try
        {
            var firstCountryHolidaysList = new List<Holiday>();
            var secondCountryHolidaysList = new List<Holiday>();

            using (var spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Stylesheet workbookstylesheet = m_PlannerStyleService.GenerateStylesheet();
                WorkbookStylesPart stylesheet = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesheet.Stylesheet = workbookstylesheet;
                stylesheet.Stylesheet.Save();

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Planner" };
                sheets.Append(sheet);

                var culture = new CultureInfo("de-DE");

                foreach (var holiday in allHolidays)
                {
                    if (holiday.CountryCode == firstCountryCode)
                    {
                        firstCountryHolidaysList.Add(holiday);
                    }
                    else if (holiday.CountryCode == secondCountryCode)
                    {
                        secondCountryHolidaysList.Add(holiday);
                    }
                }

                var columnIndex = 1;
                var mergeCells = new MergeCells();

                for (var date = from; date <= to; date = date.AddMonths(1))
                {
                    string monthName = date.ToString("MMMM yyyy", culture);

                    Cell monthYearCell = new(new CellValue(monthName))
                    {
                        DataType = CellValues.String,
                        StyleIndex = 1
                    };

                    SpreadsheetService.AppendCellToWorksheet(worksheetPart, monthYearCell, 1, (uint)columnIndex);

                    Cell emptyCell = new(new CellValue(string.Empty))
                    {
                        DataType = CellValues.String,
                        StyleIndex = 1
                    };
                    SpreadsheetService.AppendCellToWorksheet(worksheetPart, emptyCell, 1, (uint)(columnIndex + 1));

                    string cellReference1 = SpreadsheetService.GetColumnName((uint)columnIndex) + "1";
                    string cellReference2 = SpreadsheetService.GetColumnName((uint)(columnIndex + 1)) + "1";
                    MergeCell mergeCell = new MergeCell() { Reference = new StringValue($"{cellReference1}:{cellReference2}") };
                    mergeCells.Append(mergeCell);

                    SpreadsheetService.SetRowHeight(worksheetPart, 70, 1);

                    var currentRow = 2;
                    for (DateOnly currentDate = date; currentDate.Month == date.Month; currentDate = currentDate.AddDays(1))
                    {
                        string dayOfWeek = currentDate.ToString("ddd", culture);
                        string day = currentDate.Day.ToString();

                        string firstCountryHolidayName = m_HolidayNameService.GetHolidayName(currentDate, firstCountryHolidaysList);
                        string secondCountryHolidayName = m_HolidayNameService.GetHolidayName(currentDate, secondCountryHolidaysList);
                        string milestoneDescriptions = GetMilestoneDescriptionsForDate(milestones, currentDate);

                        string milestoneText = "";
                        string holidayText = "";

                        if (!string.IsNullOrEmpty(milestoneDescriptions))
                            milestoneText += $"MS: {milestoneDescriptions}";

                        if (!string.IsNullOrEmpty(firstCountryHolidayName) && !string.IsNullOrEmpty(secondCountryHolidayName))
                            holidayText += $"{firstCountryCode}&{secondCountryCode}: {firstCountryHolidayName}";
                        else if (!string.IsNullOrEmpty(firstCountryHolidayName))
                            holidayText += $"{firstCountryCode}: {firstCountryHolidayName}";
                        else if (!string.IsNullOrEmpty(secondCountryHolidayName))
                            holidayText += $"{secondCountryCode}: {secondCountryHolidayName}";

                        Cell dayCell = new Cell(new CellValue($"{day} {dayOfWeek}"))
                        {
                            DataType = CellValues.String,
                            StyleIndex = dayOfWeek == "Sa" ? 3u : dayOfWeek == "So" ? 4u : 2u
                        };

                        Cell milestoneCell = new Cell(new CellValue(milestoneText))
                        {
                            DataType = CellValues.String,
                            StyleIndex = string.IsNullOrEmpty(milestoneText) && !string.IsNullOrEmpty(firstCountryHolidayName) ? 9u : 
                            string.IsNullOrEmpty(milestoneText) && !string.IsNullOrEmpty(secondCountryHolidayName) ? 10u : string.IsNullOrEmpty(milestoneText) ? 8u : 7u
                        };

                        Cell holidayCell = new Cell(new CellValue(holidayText))
                        {
                            DataType = CellValues.String,
                            StyleIndex = !string.IsNullOrEmpty(firstCountryHolidayName) ? 5u : !string.IsNullOrEmpty(secondCountryHolidayName) ? 6u : 
                            (string.IsNullOrEmpty(firstCountryHolidayName) && !string.IsNullOrEmpty(milestoneText) ? 12u : 
                            (string.IsNullOrEmpty(firstCountryHolidayName) && !string.IsNullOrEmpty(milestoneText) ? 10u : string.IsNullOrEmpty(holidayText) ? 11u : 7u))
                        };

                        SpreadsheetService.AppendCellToWorksheet(worksheetPart, dayCell, (uint)currentRow, (uint)columnIndex);
                        SpreadsheetService.AppendCellToWorksheet(worksheetPart, milestoneCell, (uint)currentRow, (uint)(columnIndex + 1));
                        currentRow++;
                        SpreadsheetService.AppendCellToWorksheet(worksheetPart, holidayCell, (uint)currentRow, (uint)(columnIndex + 1));

                        string dayCellReference1 = SpreadsheetService.GetColumnName((uint)columnIndex) + (currentRow - 1).ToString();
                        string dayCellReference2 = SpreadsheetService.GetColumnName((uint)columnIndex) + currentRow.ToString();
                        MergeCell dayMergeCell = new MergeCell() { Reference = new StringValue($"{dayCellReference1}:{dayCellReference2}") };
                        mergeCells.Append(dayMergeCell);

                        SpreadsheetService.SetRowHeight(worksheetPart, 35, (uint)currentRow);
                        currentRow++;
                    }

                    SpreadsheetService.SetColumnWidth(worksheetPart, (uint)columnIndex, 6);
                    SpreadsheetService.SetColumnWidth(worksheetPart, (uint)(columnIndex + 1), 18);

                    columnIndex += 2;
                }

                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());

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
    private static string GetMilestoneDescriptionsForDate(IEnumerable<Milestone> milestones, DateOnly date)
    {
        var milestoneDescriptions = milestones.Where(x => date.Equals(x.Date)).Select(x => x.Description);
        return string.Join(", ", milestoneDescriptions);
    }
    #endregion private methods
}
