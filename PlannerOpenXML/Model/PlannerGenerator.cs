using PlannerOpenXML.Dialog;
using PlannerOpenXML.Model.Xlsx;
using PlannerOpenXML.Services;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;

namespace PlannerOpenXML.Model;

public class PlannerGenerator(HolidayNameService holidayNameService, IEnumerable<Holiday> allHolidays, string? countryCode1, string? countryCode2, IEnumerable<Milestone> milestones)
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
    public async Task Generate(DateOnly from, DateOnly to, string path)
    {
        ProgressDialogWindow progressDialog = new ProgressDialogWindow();
        progressDialog.Owner = Application.Current.MainWindow;
        progressDialog.Show();

        await Task.Run(() =>
        {
            try
            {
                var source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "planner.xlsx");
                File.Copy(source, path, true);

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

                    FillSheet(sheet, from, to, styles);

                    sheet.Save();
                }

                Process.Start("explorer.exe", $"/select,\"{path}\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressDialog.Dispatcher.Invoke(() => progressDialog.Close());
            }
        });
    }
    #endregion methods

    #region private methods
    private void FillSheet(Sheet sheet, DateOnly from, DateOnly to, PlannerGeneratorStyles styles)
    {
        var firstCountryHolidaysList = GetHolidays(m_AllHolidays, m_CountryCode1);
        var secondCountryHolidaysList = GetHolidays(m_AllHolidays, m_CountryCode2);
        var month = from;
        uint currentColumn = 1;
        sheet.SetRowHeight(1, styles.Row1Height);
        do
        {
            FillMonth(month, currentColumn, sheet, styles, firstCountryHolidaysList, secondCountryHolidaysList);

            month = month.AddMonths(1);
            currentColumn += 3;
        } while (month <= to);
    }

    private void FillMonth(DateOnly month, uint column, Sheet sheet, PlannerGeneratorStyles styles, IEnumerable<Holiday> holidays1, IEnumerable<Holiday> holidays2)
    {
        var date = month.ToString("MMMM yyyy", m_Ci);
        sheet.SetValue(new(column, 1), new CellSharedStringValue(date, styles.Month));
        sheet.SetValue(new(column + 1, 1), new CellEmptyValue(styles.Month));
        sheet.SetValue(new(column + 2, 1), new CellEmptyValue(styles.Month));
        sheet.Merge(new(column, 1, column + 2, 1));
        sheet.SetColumnWidth(column, styles.Column1Width);
        sheet.SetColumnWidth(column + 1, styles.Column2Width);
        sheet.SetColumnWidth(column + 2, styles.Column3Width);

        uint currentRow = 2;
        for (var day = month; day.Month == month.Month; day = day.AddDays(1))
        {
            FillDay(day, column, currentRow, sheet, styles, holidays1, holidays2);

            currentRow += 2;
        }
    }

    private void FillDay(DateOnly day, uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles, IEnumerable<Holiday> holidays1, IEnumerable<Holiday> holidays2)
    {
        // set the day row
        FillDayCell(day, column, row, sheet, styles);

        // set correct descriptions and styles for milestones and holidays
        var milestoneText = GetMilestoneDescriptionsForDate(m_Milestones, day);
        if (!string.IsNullOrEmpty(milestoneText))
            milestoneText = $"MS: {milestoneText}";
        var (holidayText, holidayStyling) = EvaluateHolidayData(day, holidays1, holidays2);

        // select correct row style
        var (styleRow1, styleRow2) = EvaluateRowStyles(milestoneText, holidayText, holidayStyling);

        // set values
        FillHolidaysAndMilestoneCells(column, row, sheet, styles, milestoneText, holidayText, styleRow1, styleRow2);

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
            sheet.SetValue(new(column + 2, row), new CellIntegerValue(weekOfYear, styles.Week));
            sheet.SetValue(new(column + 2, row + 1), new CellEmptyValue(styles.Week));
            sheet.Merge(new(column + 2, row, column + 2, row + 1));
        }
        else
        {
            sheet.SetValue(
                new(column + 2, row),
                new CellSharedStringValue(string.Empty, styles.GetStyleIndex(styleRow1, 2, 1)));

            sheet.SetValue(
                new(column + 2, row + 1),
                new CellSharedStringValue(string.Empty, styles.GetStyleIndex(styleRow2, 2, 2)));
        }
    }

    private void FillDayCell(DateOnly day, uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles)
    {
        var dayOfWeek = day.ToString("ddd", m_Ci);
        var dayStyle = day.DayOfWeek switch
        {
            DayOfWeek.Saturday => styles.Saturday,
            DayOfWeek.Sunday => styles.Sunday,
            _ => styles.Day,
        };
        sheet.SetValue(new(column, row), new CellSharedStringValue($"{day.Day} {dayOfWeek}", dayStyle));
        // notice: needs to fill also empty cells with style, otherwise excel messes up borders
        sheet.SetValue(new(column, row + 1), new CellEmptyValue(dayStyle));
        sheet.Merge(new(column, row, column, row + 1));
    }

    private static void FillHolidaysAndMilestoneCells(uint column, uint row, Sheet sheet, PlannerGeneratorStyles styles, string milestoneText, string holidayText, CellStyle styleRow1, CellStyle styleRow2)
    {
        if (!string.IsNullOrEmpty(milestoneText))
        {
            sheet.SetValue(
                new(column + 1, row),
                new CellSharedStringValue(milestoneText, styles.GetStyleIndex(styleRow1, 1, 1)));
        }
        else
        {
            sheet.SetValue(
                new(column + 1, row),
                new CellEmptyValue(styles.GetStyleIndex(styleRow1, 1, 1)));
        }

        if (!string.IsNullOrEmpty(holidayText))
        {
            sheet.SetValue(
                new(column + 1, row + 1),
                new CellSharedStringValue(holidayText, styles.GetStyleIndex(styleRow2, 1, 2)));
        }
        else
        {
            sheet.SetValue(
                new(column + 1, row + 1),
                new CellEmptyValue(styles.GetStyleIndex(styleRow2, 1, 2)));
        }
    }

    private static (CellStyle styleRow1, CellStyle styleRow2) EvaluateRowStyles(string milestoneText, string holidayText, CellStyle holidayStyling)
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