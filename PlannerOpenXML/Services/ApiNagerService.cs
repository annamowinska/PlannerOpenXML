using Newtonsoft.Json;
using PlannerOpenXML.Converters;
using PlannerOpenXML.Model;
using System.Net.Http;

namespace PlannerOpenXML.Services;

public class ApiNagerService : IApiService
{
    #region fields
    private readonly HttpClient m_HttpClient = new();
    private readonly IHolidayConverter m_HolidayConverter;
    #endregion fields

    #region constructors
    public ApiNagerService(IHolidayConverter holidayConverter)
    {
        m_HolidayConverter = holidayConverter;
    }
    #endregion constructors

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
            var nagerHolidays = JsonConvert.DeserializeObject<IEnumerable<NagerHoliday>>(json);
            if (nagerHolidays == null)
            {
                Console.WriteLine($"Could not deserialize content for {countryCode}: \"{json}\"");
                return Array.Empty<Holiday>();
            }

            var holidays = m_HolidayConverter.Convert(nagerHolidays);
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

    public async Task<IEnumerable<CountryList>> GetAvailableCountriesAsync()
    {
        try
        {
            var response = await m_HttpClient.GetAsync("https://date.nager.at/api/v3/AvailableCountries");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            var nagerCountries = JsonConvert.DeserializeObject<IEnumerable<NagerCountry>>(json);
            if (nagerCountries == null)
            {
                Console.WriteLine($"Could not deserialize content: \"{json}\"");
                return Array.Empty<CountryList>();
            }

            var countries = new List<CountryList>();
            foreach (var nagerCountry in nagerCountries)
            {
                countries.Add(new CountryList(nagerCountry.Name, nagerCountry.CountryCode));
            }

            return countries;
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"An error occurred while fetching available countries: {ex.Message}");
        }
        catch (JsonException ex)
        {
            Console.WriteLine($"An error occurred while deserializing available countries: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        return Array.Empty<CountryList>();
    }
    #endregion methods
}