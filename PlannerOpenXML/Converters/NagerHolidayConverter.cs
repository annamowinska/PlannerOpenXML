using PlannerOpenXML.Model;
using System.Globalization;

namespace PlannerOpenXML.Converters;

/// <summary>
/// Converts a collection of NagerHoliday objects to a collection of Holiday objects.
/// </summary>
public class NagerHolidayConverter : IHolidayConverter
{
    public IEnumerable<Holiday> Convert(IEnumerable<NagerHoliday> nagerHolidays)
    {
        var holidays = new List<Holiday>();
        foreach (var nagerHoliday in nagerHolidays)
        {
            DateOnly date = DateOnly.ParseExact(nagerHoliday.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

            var holiday = new Holiday
            {
                Name = nagerHoliday.Name,
                Date = date,
                CountryCode = nagerHoliday.CountryCode
            };
            holidays.Add(holiday);
        }

        return holidays;
    }
}
