using Newtonsoft.Json;
using System.Net.Http;
using PlannerOpenXML.Model;

namespace PlannerOpenXML.Services;

public class ApiService
{
    #region fields
    private readonly HttpClient m_HttpClient = new();
    #endregion fields

    #region methods
    /// <summary>
    /// Connects to a web service and collects holidays for a given year and country.
    /// </summary>
    /// <param name="year">The year</param>
    /// <param name="countryCode">The country code</param>
    /// <returns>A list of holidays</returns>
    public async Task<IEnumerable<Holiday>> GetHolidaysAsync(int year, string countryCode)
    {
        try
        {
            var response = await m_HttpClient.GetAsync($"https://date.nager.at/api/v3/PublicHolidays/{year}/{countryCode}");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            var holidays = JsonConvert.DeserializeObject<IEnumerable<Holiday>>(json);
            if (holidays == null)
            {
                Console.WriteLine($"Could not deserialize content for {countryCode}: \"{json}\"");
                return Array.Empty<Holiday>();
            }
            return holidays;
        }

        catch (HttpRequestException ex)
        {
            Console.WriteLine($"An error occurred while fetching holidays for {countryCode}: {ex.Message}");
        }

        catch (JsonException ex)
        {
            Console.WriteLine($"An error occurred while deserializing holidays for {countryCode}: {ex.Message}");
        }

        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        return Array.Empty<Holiday>();
    }

    /// <summary>
    /// Collects holidays for multiple countries for a given year.
    /// </summary>
    /// <param name="year">The year</param>
    /// <param name="countryCodes">A list of countries</param>
    /// <returns>a single list of holidays from all selected countries at once.</returns>
    public async Task<IEnumerable<Holiday>> GetHolidaysForCountriesAsync(int year, IEnumerable<string> countryCodes)
    {
        var result = new List<Holiday>();

        foreach (string countryCode in countryCodes)
        {
            result.AddRange(await GetHolidaysAsync(year, countryCode));
        }

        return result;
    }
    #endregion methods
}
