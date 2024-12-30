using Newtonsoft.Json;
using PlannerOpenXML.Model;
using System.IO;

namespace PlannerOpenXML.Services;

public class HolidayCacheService(IApiService apiService, INotificationService notificationService)
{
    #region fields
    private readonly IApiService m_ApiService = apiService;
    private readonly INotificationService m_NotificationService = notificationService;
    private readonly string m_Path
        = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Corsol GmbH",
            "Planner",
            "holiday.json");

    #endregion fields

    #region methods
    public async Task<List<Holiday>> GetAllHolidaysInRangeAsync(DateOnly fromDate, DateOnly toDate, List<string> selectedCountryCodes)
    {
        var allHolidays = new List<Holiday>();
        var yearsFromUserInput = Enumerable.Range(fromDate.Year, toDate.Year - fromDate.Year + 1).ToList();
        var countryservice = ServiceContainer.GetService<ICountryListService>();
        var countryCodes = countryservice.GetCountryCodes();
        var countryCodesMessage = string.Join(", ", countryCodes);

        if (File.Exists(m_Path))
        {
            var holidaysFromFile = ReadHolidaysFromFile().ToList();
            var holidaysGroupedByCountryAndYear = holidaysFromFile.GroupBy(h => new { h.CountryCode, h.Date.Year }).ToList();
            var yearsInFile = holidaysFromFile.Select(h => h.Date.Year).Distinct().ToList();
            var missingYears = yearsFromUserInput.Except(yearsInFile).ToList();
            var missingYearsMessage = string.Join(", ", missingYears);

            if (missingYears.Count != 0)
            {
                if (!InternetAvailabilityService.IsInternetAvailable())
                {
                    m_NotificationService.NotifyError($"No internet connection. Failed to download {missingYearsMessage} holidays. A planner will be created without {missingYearsMessage} holidays applied.");
                    return [];
                }
                else
                {
                    foreach (var missingYear in missingYears)
                    {
                        foreach (var countryCode in selectedCountryCodes)
                        {
                            var fetchedHolidays = await m_ApiService.GetHolidaysAsync(missingYear, countryCode);
                            if (countryCode == "DE")
                            {
                                fetchedHolidays = fetchedHolidays
                                    .Where(h => h.Counties == null || h.Counties.Intersect(new[] { "DE-BY", "DE-SL" }).Any())
                                    .ToList();
                            }
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
                if (holidaysFromFile.Any(h => selectedCountryCodes.Contains(h.CountryCode)))
                {
                    foreach (var countryCode in selectedCountryCodes)
                    {
                        var missingYearsForCountry = yearsFromUserInput.Except(holidaysGroupedByCountryAndYear
                            .Where(g => g.Key.CountryCode == countryCode)
                            .Select(g => g.Key.Year))
                            .ToList();
                        if (missingYearsForCountry.Count != 0)
                        {
                            if (!InternetAvailabilityService.IsInternetAvailable())
                            {
                                m_NotificationService.NotifyError($"No internet connection. Failed to download {missingYearsMessage} holidays. A planner will be created without {missingYearsMessage} holidays applied.");
                                return [];
                            }
                            else
                            {
                                foreach (var year in missingYearsForCountry)
                                {
                                    var fetchedHolidays = await m_ApiService.GetHolidaysAsync(year, countryCode);
                                    if (countryCode == "DE")
                                    {
                                        fetchedHolidays = fetchedHolidays
                                            .Where(h => h.Counties == null || h.Counties.Intersect(new[] { "DE-BY", "DE-SL" }).Any())
                                            .ToList();
                                    }
                                    allHolidays.AddRange(fetchedHolidays);
                                }
                            }
                            allHolidays.AddRange(holidaysFromFile);
                            allHolidays.Sort((x, y) => x.Date.CompareTo(y.Date));
                            SaveHolidaysToFile(allHolidays);
                        }
                        else
                        {
                            allHolidays.AddRange(holidaysFromFile);
                        }
                    }
                }
                else
                {
                    if (!InternetAvailabilityService.IsInternetAvailable())
                    {
                        m_NotificationService.NotifyError($"No internet connection. Failed to download {countryCodesMessage} holidays. A planner will be created without {countryCodesMessage} holidays applied.");
                        return [];
                    }
                    else
                    {
                        foreach (var year in yearsFromUserInput)
                        {
                            foreach (var countryCode in countryCodes)
                            {
                                var fetchedHolidays = await m_ApiService.GetHolidaysAsync(year, countryCode);
                                if (countryCode == "DE")
                                {
                                    fetchedHolidays = fetchedHolidays
                                        .Where(h => h.Counties == null || h.Counties.Intersect(new[] { "DE-BY", "DE-SL" }).Any())
                                        .ToList();
                                }
                            allHolidays.AddRange(fetchedHolidays);
                                allHolidays.AddRange(fetchedHolidays);
                            }
                            allHolidays.AddRange(holidaysFromFile);
                            allHolidays.Sort((x, y) => x.Date.CompareTo(y.Date));
                            SaveHolidaysToFile(allHolidays);
                        }
                    }
                }
            }
        }
        else
        {
            if (!InternetAvailabilityService.IsInternetAvailable())
            {
                m_NotificationService.NotifyError("No internet connection. Failed to download the holidays. An empty planner will be created, with no holidays applied.");
                return [];
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
                            if (countryCode == "DE")
                            {
                                fetchedHolidays = fetchedHolidays
                                    .Where(h => h.Counties == null || h.Counties.Intersect(new[] { "DE-BY", "DE-SL" }).Any())
                                    .ToList();
                            }
                            allHolidays.AddRange(fetchedHolidays);
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
                            if (countryCode == "DE")
                            {
                                fetchedHolidays = fetchedHolidays
                                    .Where(h => h.Counties == null || h.Counties.Intersect(new[] { "DE-BY", "DE-SL" }).Any())
                                    .ToList();
                            }
                            allHolidays.AddRange(fetchedHolidays);
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
            Directory.CreateDirectory(Path.GetDirectoryName(m_Path));
            File.WriteAllText(m_Path, json);
        }
        catch (Exception ex)
        {
            m_NotificationService.NotifyError($"Error while saving holidays to file: {ex.Message}");
            throw;
        }
    }

    private IEnumerable<Holiday> ReadHolidaysFromFile()
    {
        try
        {
            if (File.Exists(m_Path))
            {
                var json = File.ReadAllText(m_Path);
                return JsonConvert.DeserializeObject<IEnumerable<Holiday>>(json) ?? [];
            }
            else
            {
                return [];
            }
        }
        catch (Exception ex)
        {
            m_NotificationService.NotifyError($"Error while reading holidays from file: {ex.Message}");
            throw;
        }
    }
    #endregion private methods
}