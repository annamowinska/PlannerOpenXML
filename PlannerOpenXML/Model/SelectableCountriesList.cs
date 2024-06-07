using CommunityToolkit.Mvvm.ComponentModel;
using PlannerOpenXML.Services;
using System.Collections.ObjectModel;

public class SelectableCountiesList : ObservableObject
{
    #region fields
    private readonly IApiService m_ApiService;
    #endregion fields

    #region properties
    public ObservableCollection<SelectableCountry> Countries { get; } = new ObservableCollection<SelectableCountry>();
    #endregion properties

    #region constructors
    public SelectableCountiesList(IApiService apiService)
    {
        m_ApiService = apiService;
        LoadCountriesAsync();
    }
    #endregion constructors

    #region methods
    private async void LoadCountriesAsync()
    {
        var countries = await m_ApiService.GetAvailableCountriesAsync();
        foreach (var country in countries)
        {
            Countries.Add(new SelectableCountry(country.Name, country.Code));
        }
    }
    #endregion methods
}