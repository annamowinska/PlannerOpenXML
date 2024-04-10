using Newtonsoft.Json;
using System.Net.Http;
using Planner.Model;

namespace Planner.Services
{
    public class ApiService
    {
        private readonly HttpClient _httpClient;

        public ApiService()
        {
            _httpClient = new HttpClient();
        }

        //public async Task<IEnumerable<Holiday>> GetHolidaysAsync(int year, string countryCode)
        //{
        //    var response = await _httpClient.GetAsync($"https://date.nager.at/api/v3/PublicHolidays/{year}/{countryCode}");
        //    response.EnsureSuccessStatusCode();
        //    var json = await response.Content.ReadAsStringAsync();
        //    return JsonConvert.DeserializeObject<IEnumerable<Holiday>>(json);
        //}

        public async Task<IEnumerable<Holiday>> GetHolidaysAsync(int year, params string[] countryCodes)
{
    List<Holiday> allHolidays = new List<Holiday>();

    foreach (string countryCode in countryCodes)
    {
        try
        {
            var response = await _httpClient.GetAsync($"https://date.nager.at/api/v3/PublicHolidays/{year}/{countryCode}");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            var holidays = JsonConvert.DeserializeObject<IEnumerable<Holiday>>(json);
            allHolidays.AddRange(holidays);
        }
        catch (HttpRequestException ex)
        {
            // Obsługa błędu zapytania HTTP
            Console.WriteLine($"An error occurred while fetching holidays for {countryCode}: {ex.Message}");
        }
        catch (JsonException ex)
        {
            // Obsługa błędu deserializacji JSON
            Console.WriteLine($"An error occurred while deserializing holidays for {countryCode}: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Obsługa innych wyjątków
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    return allHolidays;
}
    }
}
