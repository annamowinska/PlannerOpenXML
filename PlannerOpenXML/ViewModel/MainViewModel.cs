using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using System.Windows;
using Xceed.Wpf.Toolkit;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject
{
    #region fields
    private readonly DialogService m_DialogService;
    private readonly PlannerGenerator m_PlannerGenerator;
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
    private string? m_FirstCountryCode;

    [ObservableProperty]
    private string? m_SecondCountryCode;

    [ObservableProperty]
    private bool m_FirstCountryLabelVisibility = true;

    [ObservableProperty]
    private bool m_FirstCountryComboBoxVisibility = false;

    [ObservableProperty]
    private bool m_IsFirstCountryComboBoxOpen = false;

    [ObservableProperty]
    private bool m_SecondCountryLabelVisibility = true;

    [ObservableProperty]
    private bool m_SecondCountryComboBoxVisibility = false;

    [ObservableProperty]
    private bool m_IsSecondCountryComboBoxOpen = false;
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

        await m_PlannerGenerator.GeneratePlanner(from, to, allHolidays, path, firstCountryCode, secondCountryCode, Milestones);
        
        Year = null;
        FirstMonth = null;
        NumberOfMonths = null;
        FirstCountryCode = null;
        SecondCountryCode = null;
        MonthsLabelVisibility = true;
        MonthsComboBoxVisibility = false;
        FirstCountryLabelVisibility = true;
        FirstCountryComboBoxVisibility = false;
        SecondCountryLabelVisibility = true;
        SecondCountryComboBoxVisibility = false;
        MilestonesClear();
    }

    [RelayCommand]
    private void MonthLabelClicked(string LabelName)
    {
        MonthsLabelVisibility = false;
        MonthsComboBoxVisibility = true;
        IsMonthsComboBoxOpen = true;
    }

    [RelayCommand]
    private void FirstCountryLabelClicked(string LabelName)
    {
        FirstCountryLabelVisibility = false;
        FirstCountryComboBoxVisibility = true;
        IsFirstCountryComboBoxOpen = true;
    }

    [RelayCommand]
    private void SecondCountryLabelClicked(string LabelName)
    {
        SecondCountryLabelVisibility = false;
        SecondCountryComboBoxVisibility = true;
        IsSecondCountryComboBoxOpen = true;
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
    private void ValidateCountry()
    {
        if (!string.IsNullOrEmpty(FirstCountryCode) && !CountryList.Countries.Any(c => c.Name == FirstCountryCode))
        {
            m_NotificationService.ShowNotificationCountryInput();
            FirstCountryCode = null;
        }

        if (!string.IsNullOrEmpty(SecondCountryCode) && !CountryList.Countries.Any(c => c.Name == SecondCountryCode))
        {
            m_NotificationService.ShowNotificationCountryInput();
            SecondCountryCode = null;
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
        m_PlannerGenerator = new PlannerGenerator(apiService, holidayNameService, plannerStyleService);
        m_HolidayCacheService = holidayCacheService;
        m_NotificationService = notificationService;
        m_CountryListService = countryListService;
        m_ApiService = apiService;
        Milestones = new EditableObservableCollection<Milestone>();
        CountryList = new SelectableCountriesList(apiService);
        LoadCountriesAsync().ConfigureAwait(false);
    }
    #endregion constructors
}