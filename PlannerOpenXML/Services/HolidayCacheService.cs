using Newtonsoft.Json;
using PlannerOpenXML.Model;
using System.IO;

namespace PlannerOpenXML.Services;

public class HolidayCacheService
{
    private readonly string m_FilePath;

    public HolidayCacheService(string filePath)
    {
        m_FilePath = filePath;
    }

    public async Task SaveHolidaysFromApiToFile(IApiService apiService)
    {
        var allHolidays = new List<Holiday>();

        var fromYear = 2023;
        var toYear = 2033;

        var countryCodes = new List<string> { "DE", "HU" };

        foreach (var countryCode in countryCodes)
        {
            for (var year = fromYear; year <= toYear; year++)
            {
                var fetchedHolidays = await apiService.GetHolidaysAsync(year, countryCode);
                allHolidays.AddRange(fetchedHolidays);
            }
        }

        SaveHolidaysToFile(allHolidays);
    }

    public void SaveHolidaysToFile(IEnumerable<Holiday> holidays)
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

    public IEnumerable<Holiday> ReadHolidaysFromFile()
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
}