using CommunityToolkit.Mvvm.ComponentModel;
using PlannerOpenXML.Services;
using System.Collections.ObjectModel;

public class SelectableCountriesList : ObservableObject
{
    #region fields
    private readonly IApiService m_ApiService;
    #endregion fields

    #region properties
    public ObservableCollection<CountryList> Countries { get; } = new ObservableCollection<CountryList>();
    #endregion properties

    #region constructors
    public SelectableCountriesList(IApiService apiService)
    {
        m_ApiService = apiService;
    }
    #endregion constructors

    #region methods
    public async Task LoadCountriesAsync()
    {
        Countries.Clear();
        var countries = await m_ApiService.GetAvailableCountriesAsync();
        countries = countries.OrderBy(country => country.Name).ToList();
        foreach (var country in countries)
        {
            Countries.Add(new CountryList(country.Name, country.Code));
        }
    }
    #endregion methods
}