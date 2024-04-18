using Newtonsoft.Json;
using System.Net.Http;
using PlannerOpenXML.Model;

namespace PlannerOpenXML.Services
{
    public class ApiService
    {
        private readonly HttpClient _httpClient;

        public ApiService()
        {
            _httpClient = new HttpClient();
        }
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
            }

            return allHolidays;
        }
    }
}
