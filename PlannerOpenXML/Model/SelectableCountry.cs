using CommunityToolkit.Mvvm.ComponentModel;

public class SelectableCountry : ObservableObject
{
    #region properties
    public bool IsEnabled { get; set; } = true;
    public string Name { get; set; }
    public string Code { get; set; }
    public bool IsChecked { get; set; }
    #endregion properties

    #region constructors
    public SelectableCountry(string name, string code)
    {
        Name = name;
        Code = code;
    }
    #endregion constructors
}