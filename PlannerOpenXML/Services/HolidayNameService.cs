using PlannerOpenXML.Model;
using System.Diagnostics;

namespace PlannerOpenXML.Services;

public class HolidayNameService
{
    #region methods
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

    public string GetHolidayNameForGermany(DateOnly date, IEnumerable<Holiday> holidays)
    {
        foreach (var holiday in holidays)
        {
            if (holiday.Date == date)
            {
                return holiday.LocalName;
            }
        }
        return string.Empty;
    }
    #endregion methods
}