using CommunityToolkit.Mvvm.ComponentModel;

namespace PlannerOpenXML.Model;

public partial class Milestone : ObservableObject
{
    #region properties
    [ObservableProperty]
    private string m_Description = string.Empty;

    [ObservableProperty]
    private DateOnly m_Date;
    #endregion properties
}