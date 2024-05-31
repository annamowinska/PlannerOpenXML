using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using Xceed.Wpf.Toolkit;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject, INotifyPropertyChanged
{
    #region fields
    private readonly DialogService m_DialogService;
    private readonly PlannerGenerator m_PlannerGenerator;
    private readonly HolidayCacheService m_HolidayCacheService;
    private readonly NotificationService m_NotificationService;
    private readonly ICountryListService m_CountryListService;
    private readonly MilestoneListService m_MilestoneListService;
    private readonly AddedMilestoneService m_AddedMilestoneService;
    #endregion fields

    #region properties
    /// <summary>
    /// Pregenerated list of month numbers: 1 - 12.
    /// </summary>
    public List<int> Months { get; } = Enumerable.Range(1, 12).ToList();
    public SelectableCountiesList CountryList { get; }
    public List<string> Milestones { get; } = new List<string>();

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
    private bool m_IsMonthsComboBoxOpen = false;

    [ObservableProperty]
    private string? m_FirstCountryHolidays;

    [ObservableProperty]
    private string? m_SecondCountryHolidays;

    [ObservableProperty]
    private bool m_MilestoneLabelVisibility = true;

    [ObservableProperty]
    private bool m_MilestoneComboBoxVisibility = false;

    [ObservableProperty]
    private string? m_SelectedMilestone;

    [ObservableProperty]
    private DateTime? m_SelectedMilestoneDate;

    [ObservableProperty]
    private bool m_IsMilestoneComboBoxOpen = false;

    #endregion properties

    #region commands
    [RelayCommand]
    private void CheckCountry(SelectableCountry country)
    {
        var checkedCountries = CountryList.Countries.Count(c => c.IsChecked);

        if (checkedCountries > 2)
        {
            country.IsChecked = false;
        }
        else
        {
            var selectedCountries = CountryList.Countries.Where(c => c.IsChecked).ToList();

            if (selectedCountries.Count == 2)
            {
                FirstCountryHolidays = selectedCountries[0].Code;
                SecondCountryHolidays = selectedCountries[1].Code;

                foreach (var remainingCountry in CountryList.Countries.Where(c => !c.IsChecked))
                {
                    remainingCountry.IsEnabled = false;
                }
            }
            else
            {
                foreach (var remainingCountry in CountryList.Countries)
                {
                    remainingCountry.IsEnabled = true;
                }

                if (selectedCountries.Count == 1)
                {
                    FirstCountryHolidays = selectedCountries[0].Code;
                    SecondCountryHolidays = null;
                }
                else
                {
                    FirstCountryHolidays = null;
                    SecondCountryHolidays = null;
                }
            }
        }
    }

    [RelayCommand]
    private async Task Generate()
    {
        //if user forgot to select informations inform him
        if (!Year.HasValue || !FirstMonth.HasValue || !NumberOfMonths.HasValue)
        {
            System.Windows.MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        await m_PlannerGenerator.GeneratePlanner(from, to, allHolidays, path, firstCountryCode, secondCountryCode, m_AddedMilestoneService.GetAddedMilestones());
        
        m_AddedMilestoneService.ClearAddedMilestones();

        Year = null;
        FirstMonth = null;
        NumberOfMonths = null;
        MonthsLabelVisibility = true;
        MilestoneLabelVisibility = true;
        MonthsComboBoxVisibility = false;
        MilestoneComboBoxVisibility = false;

        foreach (var country in CountryList.Countries)
        {
            country.IsChecked = false;
            country.IsEnabled = true;
        }
    }

    [RelayCommand]
    private void MonthLabelClicked(string LabelName)
    {
        MonthsLabelVisibility = false;
        MonthsComboBoxVisibility = true;
        IsMonthsComboBoxOpen = true;
    }

    [RelayCommand]
    private void MilestoneLabelClicked(string LabelName)
    {
        MilestoneLabelVisibility = false;
        MilestoneComboBoxVisibility = true;
        IsMilestoneComboBoxOpen = true;
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
    private void AddMilestone(AddedMilestone newAddedMilestone)
    {
        if (SelectedMilestoneDate.HasValue && !string.IsNullOrWhiteSpace(SelectedMilestone))
        {
            var dateText = SelectedMilestoneDate.Value.ToString("dd-MM-yyyy");
            if (!DateTime.TryParseExact(dateText, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                m_NotificationService.ShowNotificationErrorMilestoneDateInput();
                return;
            }

            var addedMilestone = new AddedMilestone
            {
                AddedMilestoneText = SelectedMilestone,
                AddedMilestoneDate = DateOnly.FromDateTime(SelectedMilestoneDate.Value),
            };

            m_AddedMilestoneService.AddMilestone(addedMilestone);
            SelectedMilestone = null;
            SelectedMilestoneDate = null;
            m_NotificationService.ShowNotificationAddedMilestone(addedMilestone.AddedMilestoneText, addedMilestone.AddedMilestoneDate.ToShortDateString());
        }
        else
        {
            m_NotificationService.ShowNotificationErrorMilestoneAndMilestoneDateInput();
        }
    }
    #endregion commands

    #region constructors
    public MainViewModel(
        IApiService apiService, 
        HolidayNameService holidayNameService,
        AddedMilestoneNameService addedMilestoneNameService,
        AddedMilestoneService addedMilestoneService,
        PlannerStyleService plannerStyleService, 
        HolidayCacheService holidayCacheService, 
        NotificationService notificationService, 
        DialogService dialogService,
        ICountryListService countryListService)
    {
        m_DialogService = dialogService;
        m_PlannerGenerator = new PlannerGenerator(apiService, holidayNameService, addedMilestoneNameService, plannerStyleService, "", "");
        m_HolidayCacheService = holidayCacheService;
        m_NotificationService = notificationService;
        m_CountryListService = countryListService;
        m_MilestoneListService = new MilestoneListService();
        m_AddedMilestoneService = addedMilestoneService;
        CountryList = new SelectableCountiesList();
        Milestones = m_MilestoneListService.LoadMilestonesFromJson();
    }
    #endregion constructors
}