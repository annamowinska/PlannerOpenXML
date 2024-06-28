﻿namespace PlannerOpenXML.Services;

public class CountryListService : ICountryListService
{
    #region fields
    public List<string> m_CountryCodes = [];
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

    public void AddCountryCode(string countryCode)
    {
        m_CountryCodes.Add(countryCode);
    }
    #endregion methods
}