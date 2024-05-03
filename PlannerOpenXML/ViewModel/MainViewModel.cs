using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using System.ComponentModel;
using System.Windows;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject, INotifyPropertyChanged
{
    #region fields
    private readonly DialogService m_DialogService;
    private readonly PlannerGenerator m_PlannerGenerator;
    private readonly HolidayCacheService m_HolidayCacheService;
    private readonly NotificationService m_NotificationService;
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

        var from = new DateOnly(Year.Value, FirstMonth.Value, 1);
        var to = from.AddMonths(NumberOfMonths.Value).AddDays(-1);
        var allHolidays = await m_HolidayCacheService.GetAllHolidaysInRangeAsync(from, to);

        await m_PlannerGenerator.GeneratePlanner(from, to, allHolidays, path);
 
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
        DialogService dialogService)
    {
        m_DialogService = dialogService;
        m_PlannerGenerator = new PlannerGenerator(apiService, holidayNameService, plannerStyleService);
        m_HolidayCacheService = holidayCacheService;
        m_NotificationService = notificationService;
    }
    #endregion constructors
}