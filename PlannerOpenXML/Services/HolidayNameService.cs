using PlannerOpenXML.Model;

namespace PlannerOpenXML.Services;

public class HolidayNameService
{
    public string GetHolidayName(DateOnly date, IEnumerable<Holiday> holidays)
    {
        foreach (var holiday in holidays)
        {
            if (holiday.Date == date)
            {
                return holiday.Name;
            }
        }
        return string.Empty;
    }
}
