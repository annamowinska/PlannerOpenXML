using PlannerOpenXML.Model;
using System.Globalization;

namespace PlannerOpenXML.Converters;

/// <summary>
/// Converts a collection of NagerHoliday objects to a collection of Holiday objects.
/// </summary>
public class NagerHolidayConverter : IHolidayConverter
{
    #region methods
    public IEnumerable<Holiday> Convert(IEnumerable<NagerHoliday> nagerHolidays)
    {
        var holidays = new List<Holiday>();
        foreach (var nagerHoliday in nagerHolidays)
        {
            var date = DateOnly.ParseExact(nagerHoliday.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

            var holiday = new Holiday
            {
                Name = nagerHoliday.Name,
                LocalName = nagerHoliday.LocalName,
                Date = date,
                CountryCode = nagerHoliday.CountryCode,
                Counties = nagerHoliday.Counties
            };
            holidays.Add(holiday);
        }

        return holidays;
    }
    #endregion methods
}
