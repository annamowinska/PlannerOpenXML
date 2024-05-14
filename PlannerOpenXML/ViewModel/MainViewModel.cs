﻿using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject, INotifyPropertyChanged
{
    #region fields
    private readonly DialogService m_DialogService;
    private readonly PlannerGenerator m_PlannerGenerator;
    private readonly HolidayCacheService m_HolidayCacheService;
    private readonly NotificationService m_NotificationService;
    private readonly ICountryListService m_CountryListService;
    #endregion fields

    #region properties
    /// <summary>
    /// Pregenerated list of month numbers: 1 - 12.
    /// </summary>
    public List<int> Months { get; } = Enumerable.Range(1, 12).ToList();
    public List<string> Countries { get; } = new List<string>
    {
        "DE", 
        "HU",
        "PL", 
        "US",
    };

    [ObservableProperty]
    private int? m_Year;

    [ObservableProperty]
    private int? m_FirstMonth;

    [ObservableProperty]
    private int? m_NumberOfMonths;

    [ObservableProperty]
    private bool m_MonthsLabelVisibility = true;

    [ObservableProperty]
    private bool m_MonthsComboBoxVisibility = false;
    
    [ObservableProperty]
    private bool m_FirstCountryHolidaysLabelVisibility = true;

    [ObservableProperty]
    private bool m_FirstCountryHolidaysComboBoxVisibility = false;

    [ObservableProperty]
    private bool m_SecondCountryHolidaysLabelVisibility = true;

    [ObservableProperty]
    private bool m_SecondCountryHolidaysComboBoxVisibility = false;

    [ObservableProperty]
    private string? m_FirstCountryHolidays;

    [ObservableProperty]
    private string? m_SecondCountryHolidays;
    #endregion properties

    #region commands
    [RelayCommand]
    private async Task Generate()
    {
        //if user forgot to select informations inform him
        if (!Year.HasValue || !FirstMonth.HasValue || !NumberOfMonths.HasValue)
        {
            MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var path = m_DialogService.SaveFileDialog(
            "Select file path to save the planner",
            "Excel files (*.xlsx)|*.xlsx",
            "Planner");
        if (path == null)
            return;

        var firstCountryCode = FirstCountryHolidays;
        var secondCountryCode = SecondCountryHolidays;

        var countryCodes = new List<string>();

        if (!string.IsNullOrEmpty(firstCountryCode))
            countryCodes.Add(firstCountryCode);
        if (!string.IsNullOrEmpty(secondCountryCode))
            countryCodes.Add(secondCountryCode);

        m_CountryListService.UpdateCountryCodes(countryCodes);

        var from = new DateOnly(Year.Value, FirstMonth.Value, 1);
        var to = from.AddMonths(NumberOfMonths.Value).AddDays(-1);
        var allHolidays = await m_HolidayCacheService.GetAllHolidaysInRangeAsync(from, to, countryCodes);
        

        await m_PlannerGenerator.GeneratePlanner(from, to, allHolidays, path, firstCountryCode, secondCountryCode);
 
        Year = null;
        FirstMonth = null;
        NumberOfMonths = null;
        FirstCountryHolidays = null;
        SecondCountryHolidays = null;
    }

    [RelayCommand]
    private void LabelClicked(string LabelName)
    {
        switch (LabelName)
        {
            case "MonthsLabel":
                MonthsLabelVisibility = false;
                MonthsComboBoxVisibility = true;
                break;
            case "FirstCountryHolidaysLabel":
                FirstCountryHolidaysLabelVisibility = false;
                FirstCountryHolidaysComboBoxVisibility = true;
                break;
            case "SecondCountryHolidaysLabel":
                SecondCountryHolidaysLabelVisibility = false;
                SecondCountryHolidaysComboBoxVisibility = true;
                break;
            default:
                break;
        }
    }

    [RelayCommand]
    private void CheckNumericInput(string input)
    {
        if (!string.IsNullOrEmpty(input) && !input.All(char.IsDigit))
        {
            m_NotificationService.ShowNotification();
        }
    }
    #endregion commands

    #region constructors
    public MainViewModel(
        IApiService apiService, 
        HolidayNameService holidayNameService, 
        PlannerStyleService plannerStyleService, 
        HolidayCacheService holidayCacheService, 
        NotificationService notificationService, 
        DialogService dialogService,
        ICountryListService countryListService)
    {
        m_DialogService = dialogService;
        m_PlannerGenerator = new PlannerGenerator(apiService, holidayNameService, plannerStyleService, "", "");
        m_HolidayCacheService = holidayCacheService;
        m_NotificationService = notificationService;
        m_CountryListService = countryListService;
    }
    #endregion constructors
}