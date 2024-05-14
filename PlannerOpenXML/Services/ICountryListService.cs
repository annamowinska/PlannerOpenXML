namespace PlannerOpenXML.Services;

public interface ICountryListService
{
    public List<string> GetCountryCodes();

    public void UpdateCountryCodes(List<string> newCountryCodes);

    void AddCountryCode(string countryCode);
}