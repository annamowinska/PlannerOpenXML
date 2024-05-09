namespace PlannerOpenXML.Services;

public class CountryListService : ICountryListService
{
    #region fields
    private List<string> m_CountryCodes = new List<string> { "DE", "HU" };
    #endregion fields

    #region methods
    public List<string> GetCountryCodes()
    {
        return m_CountryCodes;
    }

    public void UpdateCountryCodes(List<string> newCountryCodes)
    {
        m_CountryCodes = newCountryCodes;
    }
    #endregion methods
}
