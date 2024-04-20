using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Windows;
using PlannerOpenXML.Services;

namespace PlannerOpenXML.Model;

public class PlannerGenerator
{
    #region fields
    private readonly ApiService m_ApiService = new();
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

                Stylesheet workbookstylesheet = GeneratorStylesheet();

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

                    DateTime monthDate = new DateTime(currentYear, currentMonth, 1);
                    string monthName = monthDate.ToString("MMMM yyyy", culture);

                    Cell monthYearCell = new Cell(new CellValue($"{monthName}"))
                    {
                        DataType = CellValues.String,
                        StyleIndex = 1
                    };

                    AppendCellToWorksheet(spreadsheetDocument, worksheetPart, monthYearCell, 1, (uint)(i + 1));

                    int currentRow = 2;

                    DateTime currentDate = monthDate;
                    while (currentDate.Month == monthDate.Month)
                    {
                        string dayOfWeek = currentDate.ToString("ddd", culture);
                        string cellValue = $"{currentDate.Day} {dayOfWeek}";

                        string germanHolidayName = GetHolidayName(currentDate, germanHolidays);
                        string hungarianHolidayName = GetHolidayName(currentDate, hungarianHolidays);

                        DateTime nextMonth = monthDate.AddMonths(1);
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
    private static Stylesheet GeneratorStylesheet()
    {
        var workbookstylesheet = new Stylesheet();

        var font0 = new Font();

        var font1 = new Font();
        var bold = new Bold();
        font1.Append(bold);

        var font2 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "555555" } }
        };

        var font3 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "FF0000" } }
        };

        var font4 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "00008B" } }
        };

        var font5 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "006400" } }
        };

        var font6 = new Font
        {
            Bold = new Bold(),
            Color = new Color() { Rgb = new HexBinaryValue() { Value = "C71585" } }
        };

        var fonts = new Fonts(font0, font1, font2, font3, font4, font5, font6);

        Fill fill0 = new Fill();

        Fill fill1 = new Fill();
        PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFFFF" };
        patternFill1.Append(foregroundColor1);
        fill1.Append(patternFill1);

        Fill fill2 = new Fill();
        PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "DCDCDC" };
        patternFill2.Append(foregroundColor2);
        fill2.Append(patternFill2);

        Fill fill3 = new Fill();
        PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "F68072" };
        patternFill3.Append(foregroundColor3);
        fill3.Append(patternFill3);

        Fill fill4 = new Fill();
        PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "87CEFA" };
        patternFill4.Append(foregroundColor4);
        fill4.Append(patternFill4);

        Fill fill5 = new Fill();
        PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor5 = new ForegroundColor() { Rgb = "9bff66" };
        patternFill5.Append(foregroundColor5);
        fill5.Append(patternFill5);

        Fill fill6 = new Fill();
        PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor6 = new ForegroundColor() { Rgb = "FFC0CB" };
        patternFill6.Append(foregroundColor6);
        fill6.Append(patternFill6);

        Fills fills = new Fills(fill0, fill1, fill2, fill3, fill4, fill5, fill6);

        Border border0 = new(new LeftBorder(),
                                    new RightBorder(),
                                    new TopBorder(),
                                    new BottomBorder());

        Border border1 = new Border(new LeftBorder() { Style = BorderStyleValues.Thick },
                                    new RightBorder() { Style = BorderStyleValues.Thick },
                                    new TopBorder() { Style = BorderStyleValues.Thick },
                                    new BottomBorder() { Style = BorderStyleValues.Thick });

        Border border2 = new Border(new LeftBorder() { Style = BorderStyleValues.Thick },
                                    new RightBorder() { Style = BorderStyleValues.Thick },
                                    new TopBorder() { Style = BorderStyleValues.Thin },
                                    new BottomBorder() { Style = BorderStyleValues.Thin });

        Border border3 = new Border(new LeftBorder() { Style = BorderStyleValues.Thick },
                                    new RightBorder() { Style = BorderStyleValues.Thick },
                                    new TopBorder() { Style = BorderStyleValues.Thin },
                                    new BottomBorder() { Style = BorderStyleValues.Thick });

        Borders borders = new Borders();
        borders.Append(border0);
        borders.Append(border1);
        borders.Append(border2);
        borders.Append(border3);

        CellFormat defaultStyle = new CellFormat()
        {
            FormatId = 0,
            FillId = 0,
            BorderId = 0
        };

        Alignment alignment = new Alignment()
        {
            Horizontal = HorizontalAlignmentValues.Center,
            Vertical = VerticalAlignmentValues.Center
        };

        CellFormat nameOfMonthStyle = new CellFormat(alignment)
        {
            FontId = 1,
            BorderId = 1
        };

        CellFormat borderStyle = new CellFormat()
        {
            BorderId = 2
        };

        CellFormat saturdayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 2,
            FillId = 2,
        };

        CellFormat sundayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 3,
            FillId = 3,
        };

        CellFormat germanHolidayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 4,
            FillId = 4,
        };

        CellFormat hungarianHolidayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 5,
            FillId = 5
        };

        CellFormat germanAndHungarianHolidayStyle = new CellFormat()
        {
            BorderId = 2,
            FontId = 6,
            FillId = 6
        };

        CellFormat lastDayOfMonthStyle = new CellFormat()
        {
            BorderId = 3
        };

        CellFormat lastDayOfMonthAndSaturdayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 2,
            FillId = 2
        };

        CellFormat lastDayOfMonthAndSundayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 3,
            FillId = 3
        };

        CellFormat lastDayOfMonthAndGermanHolidayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 4,
            FillId = 4
        };

        CellFormat lastDayOfMonthAndHungarianHolidayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 5,
            FillId = 5
        };

        CellFormat lastDayOfMonthAndGermanAndHungarianHolidayStyle = new CellFormat()
        {
            BorderId = 3,
            FontId = 6,
            FillId = 6
        };

        CellFormats cellformats = new CellFormats();
        cellformats.Append(defaultStyle);
        cellformats.Append(nameOfMonthStyle);
        cellformats.Append(borderStyle);
        cellformats.Append(saturdayStyle);
        cellformats.Append(sundayStyle);
        cellformats.Append(germanHolidayStyle);
        cellformats.Append(hungarianHolidayStyle);
        cellformats.Append(germanAndHungarianHolidayStyle);
        cellformats.Append(lastDayOfMonthStyle);
        cellformats.Append(lastDayOfMonthAndSaturdayStyle);
        cellformats.Append(lastDayOfMonthAndSundayStyle);
        cellformats.Append(lastDayOfMonthAndGermanHolidayStyle);
        cellformats.Append(lastDayOfMonthAndHungarianHolidayStyle);
        cellformats.Append(lastDayOfMonthAndGermanAndHungarianHolidayStyle);

        workbookstylesheet.Append(fonts);
        workbookstylesheet.Append(fills);
        workbookstylesheet.Append(borders);
        workbookstylesheet.Append(cellformats);

        return workbookstylesheet;
    }

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

    private static string GetHolidayName(DateTime date, IEnumerable<Holiday> holidays)
    {
        foreach (var holiday in holidays)
        {
            var holidayDate = DateTime.Parse(holiday.Date);
            if (holidayDate.Date == date.Date)
            {
                return holiday.Name;
            }
        }
        return string.Empty;
    }
    #endregion private methods
}
