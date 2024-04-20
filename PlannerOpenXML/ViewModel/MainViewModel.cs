using System.ComponentModel;
using System.Windows;
using PlannerOpenXML.Model;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.ComponentModel;

namespace PlannerOpenXML.ViewModel;

public partial class MainViewModel : ObservableObject, INotifyPropertyChanged
{
    #region fields
    private readonly PlannerGenerator m_PlannerGenerator = new();
    #endregion fields

    #region properties
    public List<int> Months { get; } = Enumerable.Range(1, 12).ToList();

    [ObservableProperty]
    private int? m_Year;

    [ObservableProperty]
    private int? m_FirstMonth;

    [ObservableProperty]
    private int? m_NumberOfMonths;

    [ObservableProperty]
    private Visibility m_LabelVisibility = Visibility.Visible;

    [ObservableProperty]
    private Visibility m_ComboBoxVisibility = Visibility.Collapsed;
    #endregion properties

    #region commands
    [RelayCommand]
    private async Task Generate()
    {
        await m_PlannerGenerator.GeneratePlanner(Year, FirstMonth, NumberOfMonths);
        Year = null;
        FirstMonth = null;
        NumberOfMonths = null;
    }

    [RelayCommand]
    private void LabelClicked()
    {
        LabelVisibility = Visibility.Collapsed;
        ComboBoxVisibility = Visibility.Visible;
    }
    #endregion commands
}