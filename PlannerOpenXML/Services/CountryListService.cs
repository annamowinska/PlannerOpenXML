namespace PlannerOpenXML.Services;

public static class CountryListService
{
    private static List<string> _countryCodes = new List<string> { "DE", "HU" };

    public static List<string> GetCountryCodes()
    {
        return _countryCodes;
    }

    public static void UpdateCountryCodes(List<string> newCountryCodes)
    {
        _countryCodes = newCountryCodes;
    }
}
