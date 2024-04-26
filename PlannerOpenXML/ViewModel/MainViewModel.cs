using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using System.IO;
using System.ComponentModel;
using System.Windows;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject, INotifyPropertyChanged
{
    #region constructor
    public MainViewModel(IApiService apiService, HolidayNameService holidayNameService, PlannerStyleService plannerStyleService, HolidayCacheService holidayCacheService)
    {
        m_ApiService = apiService;
        m_PlannerGenerator = new PlannerGenerator(apiService, holidayNameService, plannerStyleService);
        m_HolidayCacheService = holidayCacheService;
    }
    #endregion constructor

    #region fields
    private readonly IApiService m_ApiService;
    private readonly PlannerGenerator m_PlannerGenerator;
    private readonly HolidayCacheService m_HolidayCacheService;
    #endregion fields

    #region properties
    /// <summary>
    /// Pregenerated list of month numbers: 1 - 12.
    /// </summary>
    public List<int> Months { get; } = Enumerable.Range(1, 12).ToList();

    [ObservableProperty]
    private int? m_Year;

    [ObservableProperty]
    private int? m_FirstMonth;

    [ObservableProperty]
    private int? m_NumberOfMonths;

    [ObservableProperty]
    private bool m_LabelVisibility = true;

    [ObservableProperty]
    private bool m_ComboBoxVisibility = false;
    #endregion properties

    #region commands
    [RelayCommand]
    private async Task Generate()
    {
        // if user forgot to select informations inform him
        if (!Year.HasValue || !FirstMonth.HasValue || !NumberOfMonths.HasValue)
        {
            MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        // normal workflow: get the service trough IOC container
        var dialogService = new Services.DialogService();
        var path = dialogService.SaveFileDialog(
            "Select file path to save the planner",
            "Excel files (*.xlsx)|*.xlsx",
            "Planner");
        if (path == null)
            return;

        var from = new DateOnly(Year.Value, FirstMonth.Value, 1);
        var to = from.AddMonths(NumberOfMonths.Value).AddDays(-1);

        var fromYear = 2023;
        var toYear = 2033;

        var countryCodes = new List<string> { "DE", "HU" };
        var allHolidays = new List<Holiday>();

        if (File.Exists("holiday.json"))
        {
            var holidaysFromFile = m_HolidayCacheService.ReadHolidaysFromFile().ToList();

            var yearsInFile = holidaysFromFile.Select(h => h.Date.Year).Distinct().ToList();
            var isAllYearsPresent = Enumerable.Range(from.Year, to.Year - from.Year + 1).All(year => yearsInFile.Contains(year));

            if (isAllYearsPresent)
            {
                allHolidays.AddRange(holidaysFromFile);
            }
            else
            {
                var existingHolidays = m_HolidayCacheService.ReadHolidaysFromFile().ToList();
                var yearsFromUserInput = Enumerable.Range(from.Year, to.Year - from.Year + 1).ToList();
                var missingYears = yearsFromUserInput.Except(yearsInFile).ToList();

                foreach (var missingYear in missingYears)
                {
                    foreach (var countryCode in countryCodes)
                    {
                        var fetchedHolidays = await m_ApiService.GetHolidaysAsync(missingYear, countryCode);
                        existingHolidays.AddRange(fetchedHolidays);
                    }
                }
                existingHolidays.Sort((x, y) => x.Date.CompareTo(y.Date));

                m_HolidayCacheService.SaveHolidaysToFile(existingHolidays);
            }
        }
        else
        {
            var holidaysFromFile = m_HolidayCacheService.ReadHolidaysFromFile().ToList();
            var holidaysInRange = holidaysFromFile.Where(h => h.Date.Year >= 2023 && h.Date.Year <= 2033).ToList();

            if (holidaysInRange.Count != (2033 - 2023 + 1))
            {
                var allHolidaysList = new List<Holiday>();

                for (var year = 2023; year <= 2033; year++)
                {
                    foreach (var countryCode in countryCodes)
                    {
                        var fetchedHolidays = await m_ApiService.GetHolidaysAsync(year, countryCode);
                        allHolidaysList.AddRange(fetchedHolidays);
                    }
                }

                holidaysFromFile.AddRange(allHolidaysList);
                m_HolidayCacheService.SaveHolidaysToFile(holidaysFromFile);
            }
        }

        await m_PlannerGenerator.GeneratePlanner(from, to, m_HolidayCacheService.ReadHolidaysFromFile(), path);

        // only one reference to the instance of "m_PlannerGenerator". Do we need to keep it in the class as a local variable?
        Year = null;
        FirstMonth = null;
        NumberOfMonths = null;
    }

    [RelayCommand]
    private void LabelClicked()
    {
        LabelVisibility = false;
        ComboBoxVisibility = true;
    }
    #endregion commands
}