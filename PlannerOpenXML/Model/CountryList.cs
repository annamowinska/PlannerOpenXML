using CommunityToolkit.Mvvm.ComponentModel;

namespace PlannerOpenXML.Model;

public partial class CountryList(string name, string code) : ObservableObject
{
    #region properties
    [ObservableProperty]
    public string m_Name = name;

    [ObservableProperty]
    public string m_Code = code;
    #endregion properties
}