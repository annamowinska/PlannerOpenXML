namespace PlannerOpenXML.Model;

public class Holiday
{
    #region properties
    public string Name { get; set; } = string.Empty;
    public string LocalName { get; set; } = string.Empty;
    public DateOnly Date { get; set; } = new DateOnly();
    public string CountryCode { get; set; } = string.Empty;
    public List<string> Counties { get; set; } = new List<string>();
    #endregion properties
}