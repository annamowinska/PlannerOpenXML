using CommunityToolkit.Mvvm.ComponentModel;

public class SelectableCountry : ObservableObject
{
    #region properties
    private bool m_IsEnabled = true;
    public bool IsEnabled
    {
        get { return m_IsEnabled; }
        set { SetProperty(ref m_IsEnabled, value); }
    }

    private string m_Name;
    public string Name
    {
        get { return m_Name; }
        set { SetProperty(ref m_Name, value); }
    }

    private string m_Code;
    public string Code
    {
        get { return m_Code; }
        set { SetProperty(ref m_Code, value); }
    }
    
    private bool m_IsChecked;
    public bool IsChecked
    {
        get { return m_IsChecked; }
        set { SetProperty(ref m_IsChecked, value); }
    }
    #endregion properties

    #region constructors
    public SelectableCountry(string name, string code)
    {
        Name = name;
        Code = code;
    }
    #endregion constructors
}
