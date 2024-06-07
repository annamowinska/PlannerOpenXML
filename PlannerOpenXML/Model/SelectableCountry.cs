using CommunityToolkit.Mvvm.ComponentModel;

public partial class SelectableCountry : ObservableObject
{
    #region properties
    [ObservableProperty]
    public string m_Name;

    [ObservableProperty]
    public string m_Code;
    #endregion properties

    #region constructors
    public SelectableCountry(string name, string code)
    {
        Name = name;
        Code = code;
    }
    #endregion constructors
}