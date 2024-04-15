using ClosedXML.Excel;
using Microsoft.Win32;
using Planner.Services;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Planner.Model
{
    public class PlannerGenerator
    {
        private readonly ApiService _apiService;

        public PlannerGenerator()
        {
            _apiService = new ApiService();
        }

        public async Task GeneratePlanner(int? year, int? firstMonth, int? numberOfMonths)
        {
            if (year == null || firstMonth == null || numberOfMonths == null)
            {
                MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Select file path to save the planner";
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "Planner";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    string plannerName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                    string savePath = Path.GetDirectoryName(saveFileDialog.FileName);
                    CultureInfo culture = new CultureInfo("de-DE");

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Planner");
                        worksheet.PageSetup.PaperSize = XLPaperSize.A2Paper;

                        int currentRow = 1;
                        int currentColumn = 1;

                        int currentYear = year ?? 0;
                        int currentMonth = firstMonth ?? 0;

                        var germanHolidaysTask = _apiService.GetHolidaysAsync(currentYear, "DE");
                        var hungarianHolidaysTask = _apiService.GetHolidaysAsync(currentYear, "HU");

                        var germanHolidays = await germanHolidaysTask;
                        var hungarianHolidays = await hungarianHolidaysTask;

                        for (int i = 0; i < numberOfMonths; i++)
                        {
                            while (currentMonth > 12)
                            {
                                currentMonth -= 12;
                                currentYear++;

                                germanHolidaysTask = _apiService.GetHolidaysAsync(currentYear, "DE");
                                hungarianHolidaysTask = _apiService.GetHolidaysAsync(currentYear, "HU");
                                
                                germanHolidays = await germanHolidaysTask;
                                hungarianHolidays = await hungarianHolidaysTask;
                            }

                            DateTime monthDate = new DateTime(currentYear, currentMonth, 1);
                            string monthName = monthDate.ToString("MMMM", culture);
                            string yearMonth = $"{monthName} {currentYear}";

                            var monthCell = worksheet.Cell(currentRow, currentColumn);
                            monthCell.Value = yearMonth;
                            monthCell.Style.Font.Bold = true;
                            monthCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                            DateTime currentDate = monthDate;
                            while (currentDate.Month == monthDate.Month)
                            {
                                string dayOfWeek = currentDate.ToString("ddd", culture);
                                string cellValue = $"{currentDate.Day} {dayOfWeek}";
                                var cell = worksheet.Cell(currentRow + 1, currentColumn);

                                foreach (var holiday in germanHolidays)
                                {
                                    DateTime holidayDate = DateTime.Parse(holiday.Date);
                                    if (holidayDate.Date == currentDate.Date)
                                    {
                                        if (hungarianHolidays.Any(h => DateTime.Parse(h.Date).Date == currentDate.Date))
                                        {
                                            cellValue += $"  DE & HU: {holiday.Name}";
                                            cell.Style.Fill.BackgroundColor = XLColor.LightPink;
                                            cell.Style.Font.Bold = true;
                                            cell.Style.Font.FontColor = XLColor.DarkPink;
                                        }
                                        else
                                        {
                                            cellValue += $"  DE: {holiday.Name}";
                                            cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                                            cell.Style.Font.Bold = true;
                                            cell.Style.Font.FontColor = XLColor.DarkBlue;
                                        }
                                        break;
                                    }
                                }

                                foreach (var holiday in hungarianHolidays)
                                {
                                    DateTime holidayDate = DateTime.Parse(holiday.Date);
                                    if (holidayDate.Date == currentDate.Date)
                                    {
                                        if (!germanHolidays.Any(h => DateTime.Parse(h.Date).Date == currentDate.Date))
                                        {
                                            cellValue += $"  HU: {holiday.Name}";
                                            cell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                                            cell.Style.Font.Bold = true;
                                            cell.Style.Font.FontColor = XLColor.DarkGreen;
                                        }
                                        break;
                                    }
                                }
                                
                                cell.Value = cellValue;
                                worksheet.Column(currentColumn).Width = 35;
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                                if (currentDate.DayOfWeek == DayOfWeek.Sunday)
                                {
                                    cell.Style.Fill.BackgroundColor = XLColor.LightSalmon;
                                    cell.Style.Font.Bold = true;
                                    cell.Style.Font.FontColor = XLColor.Red;
                                }

                                if (currentDate.DayOfWeek == DayOfWeek.Saturday)
                                {
                                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                                    cell.Style.Font.Bold = true;
                                    cell.Style.Font.FontColor = XLColor.DimGray;
                                }

                                currentDate = currentDate.AddDays(1);
                                currentRow++;
                            }

                            var columnRange = worksheet.Range(currentRow - DateTime.DaysInMonth(currentYear, currentMonth), currentColumn, currentRow, currentColumn);
                            columnRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                            currentMonth++;
                            currentColumn++;
                            currentRow = 1;
                        }

                        string fullPath = Path.Combine(savePath, plannerName + ".xlsx");
                        workbook.SaveAs(fullPath);

                        MessageBox.Show($"Planner has been generated and saved as: {fullPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while generating the planner: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}