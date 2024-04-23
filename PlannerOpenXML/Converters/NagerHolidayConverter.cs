using PlannerOpenXML.Model;

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
            var holiday = new Holiday
            {
                Name = nagerHoliday.Name,
                Date = nagerHoliday.Date
            };
            holidays.Add(holiday);
        }

        return holidays;
    }
}
