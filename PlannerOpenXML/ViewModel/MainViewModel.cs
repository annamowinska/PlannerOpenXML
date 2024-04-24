using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using System.ComponentModel;
using System.Windows;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject, INotifyPropertyChanged
{
    #region constructor
    public MainViewModel(IApiService apiService, HolidayNameService holidayNameService, PlannerStyleService plannerStyleService)
    {
        m_ApiService = apiService;
        m_HolidayNameService = holidayNameService;
        m_PlannerGenerator = new PlannerGenerator(apiService, holidayNameService, plannerStyleService);
    }
    #endregion constructor

    #region fields
    private readonly IApiService m_ApiService;
    private readonly HolidayNameService m_HolidayNameService;
    private readonly PlannerGenerator m_PlannerGenerator;
    private readonly PlannerStyleService plannerStyleService;
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

        // only one reference to the instance of "m_PlannerGenerator". Do we need to keep it in the class as a local variable?
        await m_PlannerGenerator.GeneratePlanner(Year.Value, FirstMonth.Value, NumberOfMonths.Value, path);
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