namespace PlannerOpenXML.Services;

public static class CountryListService
{
    #region fields
    private static List<string> m_CountryCodes = new List<string> { "DE", "HU" };
    #endregion fields

    #region methods
    public static List<string> GetCountryCodes()
    {
        return m_CountryCodes;
    }

    public static void UpdateCountryCodes(List<string> newCountryCodes)
    {
        m_CountryCodes = newCountryCodes;
    }
    #endregion methods
}
