using CommunityToolkit.Mvvm.ComponentModel;
using Newtonsoft.Json;
using PlannerOpenXML.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

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
        bool hasInternetConnection = InternetAvailabilityService.IsInternetAvailable();
        bool loadedFromLocalFile = false;

        try
        {
            var localCountries = LoadCountriesFromLocalFile();
            if (localCountries != null && localCountries.Any())
            {
                foreach (var country in localCountries)
                {
                    Countries.Add(country);
                }
                loadedFromLocalFile = true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while loading countries from local file: {ex.Message}");
        }

        if (!loadedFromLocalFile && !hasInternetConnection)
        {
            MessageBox.Show("No internet connection and no local countries list available.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!loadedFromLocalFile)
        {
            try
            {
                Countries.Clear();
                var countriesFromApi = await m_ApiService.GetAvailableCountriesAsync();
                countriesFromApi = countriesFromApi.OrderBy(country => country.Name).ToList();

                foreach (var country in countriesFromApi)
                {
                    Countries.Add(new CountryList(country.Name, country.Code));
                }

                SaveCountriesToLocalFile(Countries);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while fetching countries from API: {ex.Message}");
            }
        }
    }
    #endregion methods

    #region private methods
    private List<CountryList> LoadCountriesFromLocalFile()
    {
        List<CountryList> countries = new List<CountryList>();

        try
        {
            string filePath = "../../../Resources/countries.json";

            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                countries = JsonConvert.DeserializeObject<List<CountryList>>(json);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while loading countries from local file: {ex.Message}");
        }

        return countries;
    }

    private void SaveCountriesToLocalFile(IEnumerable<CountryList> countries)
    {
        try
        {
            string filePath = "../../../Resources/countries.json";
            string json = JsonConvert.SerializeObject(countries);
            File.WriteAllText(filePath, json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while saving countries to local file: {ex.Message}");
        }
    }
    #endregion private methods
}