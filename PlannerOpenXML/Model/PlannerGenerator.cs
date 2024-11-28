using PlannerOpenXML.Model.Xlsx;
using PlannerOpenXML.Services;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;

namespace PlannerOpenXML.Model;

public class PlannerGenerator(
    HolidayNameService holidayNameService,
    IEnumerable<Holiday> allHolidays,
    string? countryCode1,
    string? countryCode2,
    IEnumerable<Milestone> milestones)
{
    #region fields
    private readonly HolidayNameService m_HolidayNameService = holidayNameService;
    private readonly string? m_CountryCode1 = countryCode1;
    private readonly string? m_CountryCode2 = countryCode2;
    private readonly IEnumerable<Holiday> m_AllHolidays = allHolidays;
    private readonly IEnumerable<Milestone> m_Milestones = milestones;
    private readonly CultureInfo m_Ci = new("de-DE");
    #endregion fields

    #region methods
    public async Task Generate(DateOnly from, int months, string path, string yearRange)
    {
        try
        {
            var source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "planner.xlsx");
            File.Copy(source, path, true);

            await Task.Run(() =>
            {
                using (var excel = XlsxFile.Open(path))
                {
                    if (excel is null)
                    {
                        MessageBox.Show("Failed to open the XLSX file. Something went wrong.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    var styles = new PlannerGeneratorStyles(excel);
                    if (styles.Failed)
                    {
                        MessageBox.Show("Failed to load styles from the 'Template' sheet. The sheet is broken.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    if (!excel.Sheets.TryGetByName("Planner", out var sheet))
                    {
                        MessageBox.Show("Failed to find a sheet named 'Planner'.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    FillSheet(sheet, from, months, yearRange, styles);

                    sheet.Save();
                }
            });

            Process.Start("explorer.exe", $"/select,\"{path}\"");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    #endregion methods

    #region private methods
    private void FillSheet(Sheet sheet, DateOnly from, int months, string yearRange, PlannerGeneratorStyles styles)
    {
        AddHeader(sheet, 1, months, styles, yearRange);
        var firstCountryHolidaysList = GetHolidays(m_AllHolidays, m_CountryCode1);
        var secondCountryHolidaysList = GetHolidays(m_AllHolidays, m_CountryCode2);
        var month = from;
        uint currentColumn = 1;
        
        sheet.SetRowHeight(2, styles.Row2Height);

        for (uint row = 3; row <= 64; row++)
        {
            sheet.SetRowHeight(row, styles.DayRowHeight);
        }

        sheet.PreFillToRow(new(1, 1, (uint)(months * 4), 2 + (31 * 2) + 1));
        var to = from.AddMonths(months).AddDays(-1);

        do
        {
            FillMonth(month, currentColumn, sheet, styles, firstCountryHolidaysList, secondCountryHolidaysList);

            month = month.AddMonths(1);
            currentColumn += 4;
        } while (month <= to);

        AddFooter(sheet, 1, months, styles);
    }

    private void AddHeader(Sheet sheet, uint column, int months, PlannerGeneratorStyles styles, string yearRange)
    {
        sheet.SetRowHeight(1, styles.Row1Height);

        var headerText = "Der globale Partner für innovative Automatisierung.";
        var headerYear = yearRange;

        uint headerEndColumn = column + (uint)(months * 4) - 1;

        sheet.SetValue(new(column, 1), new CellEmptyValue(styles.Header));
        sheet.SetValue(new(column + 9, 1), new CellSharedStringValue(headerText, styles.Header));
        sheet.SetValue(new(headerEndColumn - 6, 1), new CellSharedStringValue(headerYear, styles.Year));

        sheet.Merge(new(column, 1, column + 8, 1));
        sheet.Merge(new(column + 9, 1, headerEndColumn - 7, 1));
        sheet.Merge(new(headerEndColumn - 6, 1, headerEndColumn, 1));
    }

    private void AddFooter(Sheet sheet, uint column, int months, PlannerGeneratorStyles styles)
    {
        sheet.SetRowHeight(65, styles.Footer0RowHeight);
        sheet.SetRowHeight(66, styles.Footer1RowHeight);
        sheet.SetRowHeight(67, styles.Footer2RowHeight);

        uint footerEndColumn = column + (uint)(months * 4) - 1;

        var germany = "GERMANY";
        var poland = "POLAND";
        var usa = "USA";
        var slovakia = "SLOVAKIA";
        var germanyAddress = "Corsol GmbH \r\nGewerbering 17\r\n84180 Loiching";
        var polandAddress = "Corol Sp. z o.o.\r\nIgnacego Daszyńskiego 198\r\n44-100 Gliwice";
        var usaAddress = "Corsol LLC\r\n16192 Coastal Highway\r\nLewes DE 19958";
        var slovakiaAddress = "Corsol s.r.o. \r\nČernyševského 3427/26\r\n851 01 Bratislava";

        uint sectionWidth = (footerEndColumn - column + 1) / 4;
        uint germanyStart = column + 2;
        uint polandStart = germanyStart + sectionWidth;
        uint usaStart = polandStart + sectionWidth;
        uint slovakiaStart = usaStart + sectionWidth;

        sheet.Merge(new(column, 66, column + 1, 66));
        sheet.Merge(new(germanyStart, 66, polandStart - 1, 66));
        sheet.Merge(new(polandStart, 66, usaStart - 1, 66));
        sheet.Merge(new(usaStart, 66, slovakiaStart - 1, 66));
        sheet.Merge(new(slovakiaStart, 66, footerEndColumn, 66));

        sheet.SetValue(new(column, 66), new CellEmptyValue(styles.Footer1));
        sheet.SetValue(new(germanyStart, 66), new CellSharedStringValue(germany, styles.Footer1));
        sheet.SetValue(new(polandStart, 66), new CellSharedStringValue(poland, styles.Footer1));
        sheet.SetValue(new(usaStart, 66), new CellSharedStringValue(usa, styles.Footer1));
        sheet.SetValue(new(slovakiaStart, 66), new CellSharedStringValue(slovakia, styles.Footer1));

        sheet.Merge(new(column, 67, column + 1, 67));
        sheet.Merge(new(germanyStart, 67, polandStart - 1, 67));
        sheet.Merge(new(polandStart, 67, usaStart - 1, 67));
        sheet.Merge(new(usaStart, 67, slovakiaStart - 1, 67));
        sheet.Merge(new(slovakiaStart, 67, footerEndColumn, 67));

        sheet.SetValue(new(column, 67), new CellEmptyValue(styles.Footer2));
        sheet.SetValue(new(germanyStart, 67), new CellSharedStringValue(germanyAddress, styles.Footer2));
        sheet.SetValue(new(polandStart, 67), new CellSharedStringValue(polandAddress, styles.Footer2));
        sheet.SetValue(new(usaStart, 67), new CellSharedStringValue(usaAddress, styles.Footer2));
        sheet.SetValue(new(slovakiaStart, 67), new CellSharedStringValue(slovakiaAddress, styles.Footer2));
    }

    private void FillMonth(DateOnly month, uint column, Sheet sheet, PlannerGeneratorStyles styles, IEnumerable<Holiday> holidays1, IEnumerable<Holiday> holidays2)
    {
        var date = month.ToString("MMMM yyyy", m_Ci);
        sheet.SetValue(new(column, 2), new CellSharedStringValue(date, styles.Month));
        sheet.SetValue(new(column + 1, 2), new CellEmptyValue(styles.Month));
        sheet.SetValue(new(column + 2, 2), new CellEmptyValue(styles.Month));
        sheet.SetValue(new(column + 3, 2), new CellEmptyValue(styles.Month));
        sheet.Merge(new(column, 2, column + 3, 2));
        sheet.SetColumnWidth(column, styles.Column1Width);
        sheet.SetColumnWidth(column + 1, styles.Column2Width);
        sheet.SetColumnWidth(column + 2, styles.Column3Width);
        sheet.SetColumnWidth(column + 3, styles.Column4Width);

        uint currentRow = 3;
        for (var day = month; day.Month == month.Month; day = day.AddDays(1))
        {
            FillDay(day, column, currentRow, sheet, styles, holidays1, holidays2);

            currentRow += 2;
        }
    }

    private void FillDay(DateOnly day, uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles, IEnumerable<Holiday> holidays1, IEnumerable<Holiday> holidays2)
    {
        // set the day of month row
        FillDayOfMonthCell(day, column, row, sheet, styles);

        // set the day of week row
        FillDayOfWeekCell(day, column, row, sheet, styles);

        // set correct descriptions and styles for milestones and holidays
        var milestoneText = GetMilestoneDescriptionsForDate(m_Milestones, day);
        if (!string.IsNullOrEmpty(milestoneText))
            milestoneText = $"MS: {milestoneText}";
        var (holidayText, holidayStyling) = EvaluateHolidayData(day, holidays1, holidays2);

        // select correct row style
        var (styleRow1, styleRow2) = EvaluateRowStyles(milestoneText, holidayText, holidayStyling, day);

        // set values
        FillHolidaysAndMilestoneCells(day, column, row, sheet, styles, milestoneText, holidayText, styleRow1, styleRow2);

        // set the week row
        FillWeekCell(day, column, row, sheet, styles, styleRow1, styleRow2);
    }

    private static void FillWeekCell(DateOnly day, uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles, CellStyle styleRow1, CellStyle styleRow2)
    {
        if (day.DayOfWeek == DayOfWeek.Monday)
        {
            var calendar = CultureInfo.InvariantCulture.Calendar;
            int weekOfYear = calendar.GetWeekOfYear(
                day.ToDateTime(TimeOnly.MinValue),
                CalendarWeekRule.FirstFourDayWeek,
                DayOfWeek.Monday);
            
             if (weekOfYear == 53)
             {
                weekOfYear = 1;
             }

            sheet.SetValue(new(column + 3, row), new CellIntegerValue(weekOfYear, styles.Week));
            sheet.SetValue(new(column + 3, row + 1), new CellEmptyValue(styles.Week));
            sheet.Merge(new(column + 3, row, column + 3, row + 1));
        }
        else
        {
            sheet.SetValue(
                new(column + 3, row),
                new CellSharedStringValue(string.Empty, styles.GetStyleIndex(styleRow1, 2, 1)));

            sheet.SetValue(
                new(column + 3, row + 1),
                new CellSharedStringValue(string.Empty, styles.GetStyleIndex(styleRow2, 2, 2)));
        }
    }

    private void FillDayOfMonthCell(DateOnly day, uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles)
    {
        var dayStyle = day.DayOfWeek switch
        {
            DayOfWeek.Saturday => styles.MonthDaySaturday,
            DayOfWeek.Sunday => styles.MonthDaySunday,
            _ => styles.MonthDay,
        };
        sheet.SetValue(new(column, row), new CellSharedStringValue($"{day.Day}", dayStyle));
        // notice: needs to fill also empty cells with style, otherwise excel messes up borders
        sheet.SetValue(new(column, row + 1), new CellEmptyValue(dayStyle));
        sheet.Merge(new(column, row, column, row + 1));
    }

    private void FillDayOfWeekCell(DateOnly day, uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles)
    {
        var dayOfWeek = day.ToString("ddd", m_Ci);
        var dayStyle = day.DayOfWeek switch
        {
            DayOfWeek.Saturday => styles.WeekDaySaturday,
            DayOfWeek.Sunday => styles.WeekDaySunday,
            _ => styles.WeekDay,
        };
        sheet.SetValue(new(column + 1, row), new CellSharedStringValue($"{dayOfWeek}", dayStyle));
        // notice: needs to fill also empty cells with style, otherwise excel messes up borders
        sheet.SetValue(new(column + 1, row + 1), new CellEmptyValue(dayStyle));
        sheet.Merge(new(column + 1, row, column + 1, row + 1));
    }

    private static void FillHolidaysAndMilestoneCells(DateOnly day, uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles, string milestoneText, string holidayText, CellStyle styleRow1, CellStyle styleRow2)
    {
        if (!string.IsNullOrEmpty(milestoneText) && string.IsNullOrEmpty(holidayText))
        {
            if (day.DayOfWeek != DayOfWeek.Monday)
            {
                sheet.Merge(new(column + 2, row, column + 3, row + 1));
            }
            else
            {
                sheet.Merge(new(column + 2, row, column + 2, row + 1));
            }
            sheet.SetValue(
                new(column + 2, row),
                new CellSharedStringValue(milestoneText, styles.GetStyleIndex(styleRow1, 1, 1)));
            sheet.SetValue(
                new(column + 2, row + 1),
                new CellEmptyValue(styles.GetStyleIndex(styleRow2, 1, 2)));
        }

        if (string.IsNullOrEmpty(milestoneText) && !string.IsNullOrEmpty(holidayText))
        {
            if (day.DayOfWeek != DayOfWeek.Monday)
            {
                sheet.Merge(new(column + 2, row, column + 3, row + 1));
            }
            else
            {
                sheet.Merge(new(column + 2, row, column + 2, row + 1));
            }

            sheet.SetValue(
                new(column + 2, row),
                new CellSharedStringValue(holidayText, styles.GetStyleIndex(styleRow2, 1, 1)));
            sheet.SetValue(
                new(column + 2, row + 1),
                new CellEmptyValue(styles.GetStyleIndex(styleRow2, 1, 2)));
        }

        if (!string.IsNullOrEmpty(milestoneText) && !string.IsNullOrEmpty(holidayText))
        {
            if (day.DayOfWeek != DayOfWeek.Monday)
            {
                sheet.Merge(new(column + 2, row, column + 3, row));
                sheet.Merge(new(column + 2, row + 1, column + 3, row + 1));
            }
            sheet.SetValue(
                new(column + 2, row),
                new CellSharedStringValue(milestoneText, styles.GetStyleIndex(styleRow1, 1, 1)));
            sheet.SetValue(
                new(column + 2, row + 1),
                new CellSharedStringValue(holidayText, styles.GetStyleIndex(styleRow2, 1, 2)));

        }

        if (string.IsNullOrEmpty(milestoneText) && string.IsNullOrEmpty(holidayText))
        {
            if (day.DayOfWeek != DayOfWeek.Monday)
            {
                sheet.Merge(new(column + 2, row, column + 3, row + 1));
            }
            else
            {
                sheet.Merge(new(column + 2, row, column + 2, row + 1));
            }

            sheet.SetValue(
                new(column + 2, row),
                new CellEmptyValue(styles.GetStyleIndex(styleRow1, 1, 1)));
            sheet.SetValue(
                new(column + 2, row + 1),
                new CellEmptyValue(styles.GetStyleIndex(styleRow2, 1, 2)));
        }
    }

    private static (CellStyle styleRow1, CellStyle styleRow2) EvaluateRowStyles(string milestoneText, string holidayText, CellStyle holidayStyling, DateOnly day)
    {
        var styleRow1 = CellStyle.Default;
        var styleRow2 = CellStyle.Default;
        if (!string.IsNullOrEmpty(milestoneText) && !string.IsNullOrEmpty(holidayText))
        {
            styleRow1 = CellStyle.Milestone;
            styleRow2 = holidayStyling;
        }
        else if (!string.IsNullOrEmpty(milestoneText))
        {
            styleRow1 = CellStyle.Milestone;
            styleRow2 = CellStyle.Milestone;
        }
        else if (!string.IsNullOrEmpty(holidayText))
        {
            styleRow1 = holidayStyling;
            styleRow2 = holidayStyling;
        }

        return (styleRow1, styleRow2);
    }

    private (string holidayText, CellStyle holidayStyling) EvaluateHolidayData(DateOnly day, IEnumerable<Holiday> holidays1, IEnumerable<Holiday> holidays2)
    {
        var holidayText = string.Empty;
        var holidayStyling = CellStyle.Default;
        var firstCountryHolidayName = m_HolidayNameService.GetHolidayName(day, holidays1);
        var secondCountryHolidayName = m_HolidayNameService.GetHolidayName(day, holidays2);
        if (!string.IsNullOrEmpty(firstCountryHolidayName) && !string.IsNullOrEmpty(secondCountryHolidayName))
        {
            holidayText = $"{m_CountryCode1}&{m_CountryCode2}: {firstCountryHolidayName}";
            holidayStyling = CellStyle.Holiday12;
        }
        else if (!string.IsNullOrEmpty(firstCountryHolidayName))
        {
            holidayText = $"{m_CountryCode1}: {firstCountryHolidayName}";
            holidayStyling = CellStyle.Holiday1;
        }
        else if (!string.IsNullOrEmpty(secondCountryHolidayName))
        {
            holidayText = $"{m_CountryCode2}: {secondCountryHolidayName}";
            holidayStyling = CellStyle.Holiday2;
        }

        return (holidayText, holidayStyling);
    }

    private static string GetMilestoneDescriptionsForDate(IEnumerable<Milestone> milestones, DateOnly date)
    {
        var milestoneDescriptions = milestones.Where(x => date.Equals(x.Date)).Select(x => x.Description);
        return string.Join(", ", milestoneDescriptions);
    }

    private static IEnumerable<Holiday> GetHolidays(IEnumerable<Holiday> allHolidays, string? countryCode)
    {
        if (string.IsNullOrEmpty(countryCode))
            yield break;

        foreach (var holiday in allHolidays)
        {
            if (holiday.CountryCode.Equals(countryCode))
                yield return holiday;
        }
    }
    #endregion private methods
}