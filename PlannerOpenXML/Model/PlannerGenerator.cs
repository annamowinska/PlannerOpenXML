﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PlannerOpenXML.Services;
using System.Globalization;
using System.Windows;

namespace PlannerOpenXML.Model;

public class PlannerGenerator
{
    #region fields
    private readonly IApiService m_ApiService;
    private readonly HolidayNameService m_HolidayNameService;
    private readonly MilestoneNameService m_MilestoneNameService;
    private readonly PlannerStyleService m_PlannerStyleService;
    private readonly string m_FirstCountryCode;
    private readonly string m_SecondCountryCode;
    private List<Milestone> m_MilestoneList;
    #endregion fields

    #region constructors
    public PlannerGenerator(
        IApiService apiService, 
        HolidayNameService holidayNameService,
        MilestoneNameService milestoneNameService,
        PlannerStyleService plannerStyleService,
        string firstCountryCode, 
        string secondCountryCode, 
        List<Milestone> milestones)
    {
        m_ApiService = apiService;
        m_HolidayNameService = holidayNameService;
        m_MilestoneNameService = milestoneNameService;
        m_PlannerStyleService = plannerStyleService;
        m_FirstCountryCode = firstCountryCode;
        m_SecondCountryCode = secondCountryCode;
        m_MilestoneList = milestones;
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
    public async Task GeneratePlanner(DateOnly from, DateOnly to, IEnumerable<Holiday> allHolidays, string path, string firstCountryCode, string secondCountryCode)
    {
        try
        {
            List<Holiday> firstCountryHolidaysList = new List<Holiday>();
            List<Holiday> secondCountryHolidaysList = new List<Holiday>();

            var milestoneListReader = new MilestoneListReader();
            List<Milestone> milestones = milestoneListReader.LoadMilestones();

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

                CultureInfo culture = new CultureInfo("de-DE");

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

                for (DateOnly date = from; date <= to; date = date.AddMonths(1))
                {
                    string monthName = date.ToString("MMMM yyyy", culture);

                    Cell monthYearCell = new Cell(new CellValue(monthName))
                    {
                        DataType = CellValues.String,
                        StyleIndex = 1
                    };

                    SpreadsheetService.AppendCellToWorksheet(worksheetPart, monthYearCell, 1, (uint)columnIndex);

                    Cell emptyCell = new Cell(new CellValue(string.Empty))
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
                        string milestoneName = m_MilestoneNameService.GetMilestoneName(currentDate, milestones);

                        DateOnly nextMonth = date.AddMonths(1);
                        bool isLastDayOfMonth = currentDate.AddDays(1).Month != nextMonth.Month;

                        string milestoneText = "";
                        string holidayText = "";

                        if (!string.IsNullOrEmpty(milestoneName))
                            milestoneText += $" MS: {milestoneName}";

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
                        Text text = new Text($"{milestoneText}\n{holidayText}");

                        inlineString.AppendChild(text);

                        additionalInfoCell.InlineString = inlineString;
                        additionalInfoCell.DataType = CellValues.InlineString;

                        additionalInfoCell.StyleIndex = 5;

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
}