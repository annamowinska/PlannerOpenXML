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
    private bool m_LabelVisibility = true;

    [ObservableProperty]
    private bool m_ComboBoxVisibility = false;
    #endregion properties

    #region commands
    [RelayCommand]
    private async Task Generate()
    {
        if (!Year.HasValue || !FirstMonth.HasValue || !NumberOfMonths.HasValue)
        {
            MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        await m_PlannerGenerator.GeneratePlanner(Year.Value, FirstMonth.Value, NumberOfMonths.Value);
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