﻿namespace PlannerOpenXML.Services;

public interface ICountryListService
{
    #region methods
    List<string> GetCountryCodes();

    void UpdateCountryCodes(List<string> newCountryCodes);

    void AddCountryCode(string countryCode);
    #endregion methods
}