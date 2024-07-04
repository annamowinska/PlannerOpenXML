using Newtonsoft.Json;
using PlannerOpenXML.Converters;
using PlannerOpenXML.Model;
using System.Net.Http;

namespace PlannerOpenXML.Services;

public class ApiNagerService(IHolidayConverter holidayConverter, INotificationService notificationService) : IApiService
{
    #region fields
    private readonly HttpClient m_HttpClient = new();
    private readonly IHolidayConverter m_HolidayConverter = holidayConverter;
    private readonly INotificationService m_NotificationService = notificationService;
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
            var nagerHolidays = JsonConvert.DeserializeObject<IEnumerable<NagerHoliday>>(json);
            if (nagerHolidays == null)
            {
                m_NotificationService.NotifyError($"Could not deserialize content for {countryCode}: \"{json}\"");
                return [];
            }

            var holidays = m_HolidayConverter.Convert(nagerHolidays);
            return holidays;
        }

        catch (HttpRequestException ex)
        {
            m_NotificationService.NotifyError($"An error occurred while fetching holidays for {countryCode}: {ex.Message}");
        }

        catch (JsonException ex)
        {
            m_NotificationService.NotifyError($"An error occurred while deserializing holidays for {countryCode}: {ex.Message}");
        }

        catch (Exception ex)
        {
            m_NotificationService.NotifyError($"An error occurred: {ex.Message}");
        }

        return [];
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
                m_NotificationService.NotifyError($"Could not deserialize content: \"{json}\"");
                return [];
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
            m_NotificationService.NotifyError($"An error occurred while fetching available countries: {ex.Message}");
        }
        catch (JsonException ex)
        {
            m_NotificationService.NotifyError($"An error occurred while deserializing available countries: {ex.Message}");
        }
        catch (Exception ex)
        {
            m_NotificationService.NotifyError($"An error occurred: {ex.Message}");
        }

        return [];
    }
    #endregion methods
}