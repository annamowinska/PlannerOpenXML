using PlannerOpenXML.Model;

namespace PlannerOpenXML.Services;

/// <summary>
/// Provides methods for retrieving holidays information from an external API.
/// </summary>
public interface IApiService
{
    #region methods
    Task<IEnumerable<Holiday>> GetHolidaysAsync(int year, string countryCode);
    Task<IEnumerable<CountryList>> GetAvailableCountriesAsync();
    #endregion methods
}