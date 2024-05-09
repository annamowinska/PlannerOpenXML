using Newtonsoft.Json;
using PlannerOpenXML.Model;
using System.IO;
using System.Windows;

namespace PlannerOpenXML.Services;

public class HolidayCacheService
{
    #region fields
    private readonly string m_FilePath;
    private readonly IApiService m_ApiService;
    #endregion fields

    #region constructor
    public HolidayCacheService(IApiService apiService, string filePath = "holiday.json")
    {
        m_FilePath = filePath;
        m_ApiService = apiService;
    }
    #endregion constructor

    #region methods
    public async Task<List<Holiday>> GetAllHolidaysInRangeAsync(DateOnly fromDate, DateOnly toDate)
    {
        var allHolidays = new List<Holiday>();
        var yearsFromUserInput = Enumerable.Range(fromDate.Year, toDate.Year - fromDate.Year + 1).ToList();
        var countryservice = ServiceContainer.GetService<ICountryListService>();
        var countryCodes = countryservice.GetCountryCodes();

        if (File.Exists(m_FilePath))
        {
            var holidaysFromFile = ReadHolidaysFromFile().ToList();
            var yearsInFile = holidaysFromFile.Select(h => h.Date.Year).Distinct().ToList();
            var missingYears = yearsFromUserInput.Except(yearsInFile).ToList();
            var missingYearsMessage = string.Join(", ", missingYears);

            if (missingYears.Any())
            {
                if (!InternetAvailabilityService.IsInternetAvailable())
                {
                    MessageBox.Show($"No internet connection. Failed to download {missingYearsMessage} holidays. A planner will be created without {missingYearsMessage} holidays applied.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return new List<Holiday>();
                }
                else
                {
                    foreach (var missingYear in missingYears)
                    {
                        foreach (var countryCode in countryCodes)
                        {
                            var fetchedHolidays = await m_ApiService.GetHolidaysAsync(missingYear, countryCode);
                            allHolidays.AddRange(fetchedHolidays);
                        }
                    }
                    allHolidays.AddRange(holidaysFromFile);
                    allHolidays.Sort((x, y) => x.Date.CompareTo(y.Date));
                    SaveHolidaysToFile(allHolidays);
                }
            }
            else
            {
                allHolidays.AddRange(holidaysFromFile);
            }
        }
        else
        {
            if (!InternetAvailabilityService.IsInternetAvailable())
            {
                MessageBox.Show("No internet connection. Failed to download the holidays. An empty planner will be created, with no holidays applied.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<Holiday>();
            }
            else
            {
                if (!Enumerable.Range(2023, 2033 - 2023 + 1).Intersect(yearsFromUserInput).Any())
                {
                    foreach (var year in Enumerable.Range(2023, 2033 - 2023 + 1).Union(yearsFromUserInput))
                    {
                        foreach (var countryCode in countryCodes)
                        {
                            var fetchedHolidays = await m_ApiService.GetHolidaysAsync(year, countryCode);
                            allHolidays.AddRange(fetchedHolidays);
                        }
                    }
                }
                else
                {
                    for (var year = 2023; year <= 2033; year++)
                    {
                        foreach (var countryCode in countryCodes)
                        {
                            var fetchedHolidays = await m_ApiService.GetHolidaysAsync(year, countryCode);
                            allHolidays.AddRange(fetchedHolidays);
                        }
                    }
                }
            }

            allHolidays.Sort((x, y) => x.Date.CompareTo(y.Date));
            SaveHolidaysToFile(allHolidays);
        }

        return allHolidays;
    }
    #endregion methods

    #region private methods
    private void SaveHolidaysToFile(IEnumerable<Holiday> holidays)
    {
        try
        {
            var json = JsonConvert.SerializeObject(holidays);
            File.WriteAllText(m_FilePath, json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while saving holidays to file: {ex.Message}");
            throw;
        }
    }

    private IEnumerable<Holiday> ReadHolidaysFromFile()
    {
        try
        {
            if (File.Exists(m_FilePath))
            {
                var json = File.ReadAllText(m_FilePath);
                return JsonConvert.DeserializeObject<IEnumerable<Holiday>>(json);
            }
            else
            {
                return new List<Holiday>();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while reading holidays from file: {ex.Message}");
            throw;
        }
    }
    #endregion private methods
}