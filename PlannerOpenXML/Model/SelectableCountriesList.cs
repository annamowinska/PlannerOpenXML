using CommunityToolkit.Mvvm.ComponentModel;
using PlannerOpenXML.Services;
using System.Collections.ObjectModel;

public class SelectableCountiesList : ObservableObject
{
    #region fields
    private readonly IApiService m_ApiService;
    #endregion fields

    #region properties
    public ObservableCollection<CountryList> Countries { get; } = new ObservableCollection<CountryList>();
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
        countries = countries.OrderBy(country => country.Name).ToList();
        foreach (var country in countries)
        {
            Countries.Add(new CountryList(country.Name, country.Code));
        }
    }
    #endregion methods
}