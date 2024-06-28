﻿using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using System.Windows;
using Xceed.Wpf.Toolkit;
using System.Windows.Controls;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject
{
    #region fields
    private readonly DialogService m_DialogService;
    private readonly PlannerGenerator_old m_PlannerGenerator;
    private readonly HolidayNameService m_HolidayNameService;
    private readonly HolidayCacheService m_HolidayCacheService;
    private readonly NotificationService m_NotificationService;
    private readonly ICountryListService m_CountryListService;
    private readonly IApiService m_ApiService;
    #endregion fields

    #region properties
    /// <summary>
    /// Pregenerated list of month numbers: 1 - 12.
    /// </summary>
    public List<int> Months { get; } = Enumerable.Range(1, 12).ToList();
    public SelectableCountriesList CountryList { get; }

    [ObservableProperty]
    private EditableObservableCollection<Milestone> m_Milestones;

    [ObservableProperty]
    private int? m_Year = DateTime.Now.Year;

    [ObservableProperty]
    private int? m_FirstMonth = 1;

    [ObservableProperty]
    private int? m_NumberOfMonths = 12;

    [ObservableProperty]
    private string? m_FirstCountryCode;

    [ObservableProperty]
    private string? m_SecondCountryCode;
    #endregion properties

    #region commands
    [RelayCommand]
    private async Task Generate()
    {
        //if user forgot to select informations inform him
        if (!Year.HasValue || !FirstMonth.HasValue || !NumberOfMonths.HasValue)
        {
            System.Windows.MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var path = m_DialogService.SaveFileWithExtensionList(
            "Select file path to save the planner",
            DialogService.XLSX_FILES);
        if (path == null)
            return;

        var firstCountryCode = FirstCountryCode;
        var secondCountryCode = SecondCountryCode;
        var countryCodes = new List<string>();

        if (!string.IsNullOrEmpty(firstCountryCode))
            countryCodes.Add(firstCountryCode);
        if (!string.IsNullOrEmpty(secondCountryCode))
            countryCodes.Add(secondCountryCode);

        m_CountryListService.UpdateCountryCodes(countryCodes);

        var from = new DateOnly(Year.Value, FirstMonth.Value, 1);
        var to = from.AddMonths(NumberOfMonths.Value).AddDays(-1);
        var allHolidays = await m_HolidayCacheService.GetAllHolidaysInRangeAsync(from, to, countryCodes);


        var generator = new PlannerGenerator(m_HolidayNameService, allHolidays, firstCountryCode, secondCountryCode, Milestones);
        generator.Generate(from, to, path);
    }

    [RelayCommand]
    private void CheckNumericInput(object parameter)
    {
        if (parameter is WatermarkTextBox textBox)
        {
            textBox.PreviewTextInput += (sender, e) =>
            {
                if (!char.IsDigit(e.Text, 0))
                {
                    e.Handled = true;
                    m_NotificationService.ShowNotificationErrorYearAndFirstMonthInput();
                }
            };
        }
    }
    
    [RelayCommand]
    private void CheckIfCountriesAreSame()
    {
        if (FirstCountryCode != null && SecondCountryCode != null && FirstCountryCode.Equals(SecondCountryCode))
        {
            m_NotificationService.ShowNotificationIsSameCountriesSelected();
        }
    }

    [RelayCommand]
    private void ValidateCountry(object parameter)
    {
      if (parameter is ComboBox comboBox)
      {
        var binding = comboBox.GetBindingExpression(ComboBox.SelectedValueProperty);
        if (binding != null)
        {
          binding.UpdateSource();
        }

        var countryCode = comboBox.SelectedValue?.ToString();
        var countryName = comboBox.Text;

        if (!string.IsNullOrEmpty(countryCode) && !CountryList.Countries.Any(c => c.Code == countryCode))
        {
          m_NotificationService.ShowNotificationCountryInput();
          comboBox.Text = "";
          return;
        }

        if (!string.IsNullOrEmpty(countryName) && !CountryList.Countries.Any(c => c.Name == countryName))
        {
          m_NotificationService.ShowNotificationCountryInput();
          comboBox.Text = "";
        }
      }
    }

    [RelayCommand]
    private void MilestonesLoad()
    {
        var path = m_DialogService.OpenSingleFileWithExtensionList(
            "Select file path to save the milestones",
            DialogService.JSON_FILES);
        if (path is null)
            return;

        Milestones = EditableObservableCollection<Milestone>.Load(path);
    }

    [RelayCommand]
    private void MilestonesSave()
    {
        Milestones.Save();
    }

    [RelayCommand]
    private void MilestonesSaveAs()
    {
        var path = m_DialogService.SaveFileWithExtensionList(
            "Select file path to save the milestones",
            DialogService.JSON_FILES,
            true,
            Milestones.LastPath);
        if (path is null)
            return;

        Milestones.SaveAs(path);
    }

    [RelayCommand]
    private void MilestonesClear()
    {
        Milestones.Clear();
    }

    [RelayCommand]
    private async Task LoadCountriesAsync()
    {
        await CountryList.LoadCountriesAsync();
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
        m_PlannerGenerator = new PlannerGenerator_old(apiService, holidayNameService, plannerStyleService);
        m_HolidayNameService = holidayNameService;
        m_HolidayCacheService = holidayCacheService;
        m_NotificationService = notificationService;
        m_CountryListService = countryListService;
        m_ApiService = apiService;
        Milestones = new EditableObservableCollection<Milestone>();
        CountryList = new SelectableCountriesList(apiService);
    }
    #endregion constructors
}