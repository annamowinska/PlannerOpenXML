﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using PlannerOpenXML.Services;
using System.Globalization;
using System.Windows;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;
using SpreadsheetRun = DocumentFormat.OpenXml.Spreadsheet.Run;
using SpreadsheetBold = DocumentFormat.OpenXml.Spreadsheet.Bold;
using SpreadsheetColor = DocumentFormat.OpenXml.Spreadsheet.Color;
using SpreadsheetFontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;
using SpreadsheetRunProperties = DocumentFormat.OpenXml.Spreadsheet.RunProperties;

namespace PlannerOpenXML.Model;

public class PlannerGenerator
{
    #region fields
    private readonly IApiService m_ApiService;
    private readonly HolidayNameService m_HolidayNameService;
    private readonly PlannerStyleService m_PlannerStyleService;
    private readonly string m_FirstCountryCode;
    private readonly string m_SecondCountryCode;
    #endregion fields

    #region constructors
    public PlannerGenerator(
        IApiService apiService, 
        HolidayNameService holidayNameService,
        PlannerStyleService plannerStyleService,
        string firstCountryCode, 
        string secondCountryCode)
    {
        m_ApiService = apiService;
        m_HolidayNameService = holidayNameService;
        m_PlannerStyleService = plannerStyleService;
        m_FirstCountryCode = firstCountryCode;
        m_SecondCountryCode = secondCountryCode;
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

                    var currentRow = 2;
                    for (DateOnly currentDate = date; currentDate.Month == date.Month; currentDate = currentDate.AddDays(1))
                    {
                        string dayOfWeek = currentDate.ToString("ddd", culture);
                        string day = currentDate.Day.ToString();

                        string firstCountryHolidayName = m_HolidayNameService.GetHolidayName(currentDate, firstCountryHolidaysList);
                        string secondCountryHolidayName = m_HolidayNameService.GetHolidayName(currentDate, secondCountryHolidaysList);
                        string milestoneDecsriptions = GetMilestoneDescriptionsForDate(milestones, currentDate);

                        DateOnly nextMonth = date.AddMonths(1);
                        bool isLastDayOfMonth = currentDate.AddDays(1).Month != nextMonth.Month;

                        string milestoneText = "";
                        string holidayText = "";

                        if (!string.IsNullOrEmpty(milestoneDecsriptions))
                            milestoneText += $" MS: {milestoneDecsriptions}";

                        if (!string.IsNullOrEmpty(firstCountryHolidayName) && !string.IsNullOrEmpty(secondCountryHolidayName))
                            holidayText += $" {firstCountryCode}&{secondCountryCode}: {firstCountryHolidayName}";
                        else if (!string.IsNullOrEmpty(firstCountryHolidayName))
                            holidayText += $" {firstCountryCode}: {firstCountryHolidayName}";
                        else if (!string.IsNullOrEmpty(secondCountryHolidayName))
                            holidayText += $" {secondCountryCode}: {secondCountryHolidayName}";

                        Cell dayCell = new Cell(new CellValue($"{day} {dayOfWeek}"))
                        {
                            DataType = CellValues.String,
                            StyleIndex = dayOfWeek == "Sa" ? 3u : dayOfWeek == "So" ? 4u : 2u
                        };

                        Cell additionalInfoCell = new Cell();
                        InlineString inlineString = new InlineString();

                        string[] milestoneParts = milestoneText.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                        string[] holidayParts = holidayText.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

                        foreach (var part in milestoneParts)
                        {
                            SpreadsheetText milestonePart = new SpreadsheetText(part);
                            SpreadsheetRun milestoneRun = new SpreadsheetRun(new SpreadsheetRunProperties(new SpreadsheetBold(), new SpreadsheetColor() { Rgb = new HexBinaryValue() { Value = "FF0000" } }, new SpreadsheetFontSize() { Val = 15 }), milestonePart);
                            inlineString.AppendChild(milestoneRun);

                            Paragraph milestoneParagraph = new Paragraph();
                            milestoneParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("\n"))); // Nowa linia
                            inlineString.AppendChild(milestoneParagraph);
                            
                        }

                        foreach (var part in holidayParts)
                        {
                            SpreadsheetText holidayPart = new SpreadsheetText(part);
                            SpreadsheetRun holidayRun = new SpreadsheetRun(new SpreadsheetRunProperties(new SpreadsheetBold(), new SpreadsheetColor() { Rgb = new HexBinaryValue() { Value = "009900" } }, new SpreadsheetFontSize() { Val = 10 }), holidayPart);
                            inlineString.AppendChild(holidayRun);
                        }

                        additionalInfoCell.InlineString = inlineString;
                        additionalInfoCell.DataType = CellValues.InlineString;

                        if (milestoneText.Contains("MS"))
                        {
                            additionalInfoCell.StyleIndex = 6;
                        }
                        else
                        {
                            additionalInfoCell.StyleIndex = 5;
                        }

                        SpreadsheetService.AppendCellToWorksheet(worksheetPart, dayCell, (uint)currentRow, (uint)columnIndex);
                        SpreadsheetService.AppendCellToWorksheet(worksheetPart, additionalInfoCell, (uint)currentRow, (uint)(columnIndex + 1));

                        SpreadsheetService.SetRowHeight(worksheetPart, 70);
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