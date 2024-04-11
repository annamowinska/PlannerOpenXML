using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using ClosedXML.Excel;
using System.Globalization;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Planner.Services;
using System.Windows.Controls;
using Planner.Model;
using System.Windows.Input;
using Microsoft.Toolkit.Mvvm.ComponentModel;

namespace Planner.ViewModel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private RelayCommand _generateCommand;
        public RelayCommand GenerateCommand
        {
            get
            {
                if (_generateCommand == null)
                {
                    _generateCommand = new RelayCommand(GeneratePlanner);
                }
                return _generateCommand;
            }
        }

        private int? _year;
        public int? Year
        {
            get => _year;
            set
            {
                if (_year != value)
                {
                    _year = value;
                }
            }
        }

        private int? _firstMonth;
        public int? FirstMonth
        {
            get => _firstMonth; 
            set
            {
                if (_firstMonth != value)
                {
                    _firstMonth = value;
                }
            }
        }

        private int? _numberOfMonths;

        public int? NumberOfMonths
        {
            get => _numberOfMonths;
            set
            {
                if (_numberOfMonths != value)
                {
                    _numberOfMonths = value;
                }
            }
        }

        private List<int> _months;
        public List<int> Months
        {
            get => _months;
            set
            {
                if (_months != value)
                {
                    _months = value;
                }
            }
        }

        private Visibility _labelVisibility = Visibility.Visible;
        public Visibility LabelVisibility
        {
            get => _labelVisibility;
            set
            {
                if (_labelVisibility != value)
                {
                    _labelVisibility = value;
                    OnPropertyChanged(nameof(LabelVisibility));
                }
            }
        }

        private Visibility _comboBoxVisibility = Visibility.Collapsed;
        public Visibility ComboBoxVisibility
        {
            get => _comboBoxVisibility;
            set
            {
                if (_comboBoxVisibility != value)
                {
                    _comboBoxVisibility = value;
                    OnPropertyChanged(nameof(ComboBoxVisibility));
                }
            }
        }

        private RelayCommand _labelClickedCommand;
        public RelayCommand LabelClickedCommand
        {
            get
            {
                if (_labelClickedCommand == null)
                {
                    _labelClickedCommand = new RelayCommand(LabelClicked);
                }
                return _labelClickedCommand;
            }
        }
        public void LabelClicked()
        {
            LabelVisibility = Visibility.Collapsed;
            ComboBoxVisibility = Visibility.Visible;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private readonly ApiService _apiService;

        public MainViewModel()
        {
            _apiService = new ApiService();
            InitializeMonths();
        }

        private void InitializeMonths()
        {
            Months = Enumerable.Range(1, 12).ToList();
        }

        public async void GeneratePlanner()
        {
            if (Year == null || FirstMonth == null || NumberOfMonths == null)
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

                        int currentYear = Year ?? 0;
                        int currentMonth = FirstMonth ?? 0;

                        var germanHolidaysTask = _apiService.GetHolidaysAsync(currentYear, "DE");
                        var hungarianHolidaysTask = _apiService.GetHolidaysAsync(currentYear, "HU");

                        var germanHolidays = await germanHolidaysTask;
                        var hungarianHolidays = await hungarianHolidaysTask;

                        for (int i = 0; i < NumberOfMonths; i++)
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

                        Year = null;
                        OnPropertyChanged(nameof(Year));
                        FirstMonth = null;
                        OnPropertyChanged(nameof(FirstMonth));
                        NumberOfMonths = null;
                        OnPropertyChanged(nameof(NumberOfMonths));
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