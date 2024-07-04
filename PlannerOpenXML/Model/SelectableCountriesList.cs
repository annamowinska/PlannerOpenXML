using Newtonsoft.Json;
using PlannerOpenXML.Services;
using System.Collections.ObjectModel;
using System.IO;

namespace PlannerOpenXML.Model;

public class SelectableCountriesList
{
    #region fields
    private readonly IApiService m_ApiService;
    private readonly INotificationService m_NotificationService;
    private readonly string m_Path
        = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Corsol GmbH",
            "Planner",
            "countries.json");
    #endregion fields

    #region properties
    public ObservableCollection<CountryList> Countries { get; } = [];
    #endregion properties

    #region constructors
    public SelectableCountriesList()
    {
        m_ApiService = ServiceContainer.GetService<IApiService>();
        m_NotificationService = ServiceContainer.GetService<INotificationService>();
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
            m_NotificationService.NotifyError($"An error occurred while loading countries from local file: {ex.Message}");
        }

        if (!loadedFromLocalFile && !hasInternetConnection)
        {
            m_NotificationService.NotifyError("No internet connection and no local countries list available.");
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
                m_NotificationService.NotifyError($"An error occurred while fetching countries from API: {ex.Message}");
            }
        }
    }
    #endregion methods

    #region private methods
    private List<CountryList> LoadCountriesFromLocalFile()
    {
        try
        {
            if (File.Exists(m_Path))
            {
                string json = File.ReadAllText(m_Path);
                return JsonConvert.DeserializeObject<List<CountryList>>(json) ?? [];
            }
        }
        catch (Exception ex)
        {
            m_NotificationService.NotifyError($"An error occurred while loading countries from local file: {ex.Message}");
        }

        return [];
    }

    private void SaveCountriesToLocalFile(IEnumerable<CountryList> countries)
    {
        try
        {
            string json = JsonConvert.SerializeObject(countries);
            Directory.CreateDirectory(Path.GetDirectoryName(m_Path));
            File.WriteAllText(m_Path, json);
        }
        catch (Exception ex)
        {
            m_NotificationService.NotifyError($"An error occurred while saving countries to local file: {ex.Message}");
        }
    }
    #endregion private methods
}