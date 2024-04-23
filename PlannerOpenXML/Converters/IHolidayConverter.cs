using PlannerOpenXML.Model;

namespace PlannerOpenXML.Converters;

/// <summary>
/// Converts a collection of Nager holidays to a collection of custom Holiday objects.
/// </summary>
public interface IHolidayConverter
{
    IEnumerable<Holiday> Convert(IEnumerable<NagerHoliday> nagerHolidays);
}

