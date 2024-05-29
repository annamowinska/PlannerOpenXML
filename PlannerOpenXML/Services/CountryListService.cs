namespace PlannerOpenXML.Services;

public class CountryListService : ICountryListService
{
    #region fields
    public List<string> m_CountryCodes = new List<string>();
    #endregion fields

    #region methods
    public List<string> GetCountryCodes()
    {
        return m_CountryCodes;
    }

    // Implementacja metody z interfejsu ICountryListService
    public void UpdateCountryCodes(List<string> newCountryCodes)
    {
        m_CountryCodes = newCountryCodes;
    }

    // Metoda do dodawania kodu kraju do listy
    public void AddCountryCode(string countryCode)
    {
        m_CountryCodes.Add(countryCode);
    }
    #endregion methods
}