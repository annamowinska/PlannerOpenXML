using CommunityToolkit.Mvvm.ComponentModel;

namespace PlannerOpenXML.Model;

public class SelectableCountry : ObservableObject
{
    #region properties
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
