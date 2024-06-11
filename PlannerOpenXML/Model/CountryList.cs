using CommunityToolkit.Mvvm.ComponentModel;

public partial class CountryList : ObservableObject
{
    #region properties
    [ObservableProperty]
    public string m_Name;

    [ObservableProperty]
    public string m_Code;
    #endregion properties

    #region constructors
    public CountryList(string name, string code)
    {
        Name = name;
        Code = code;
    }
    #endregion constructors
}