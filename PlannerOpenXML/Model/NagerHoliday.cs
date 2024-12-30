namespace PlannerOpenXML.Model;

public class NagerHoliday
{
    #region properties
    public string Name { get; set; } = string.Empty;
    public string LocalName { get; set; } = string.Empty;
    public string Date { get; set; } = string.Empty;
    public string CountryCode { get; set; } = string.Empty;
    public List<string> Counties { get; set; } = new List<string>();
    #endregion properties
}