namespace PlannerOpenXML.Model;

public class NagerHoliday
{
    #region properties
    /// <summary>
    /// Name of the holiday
    /// </summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>
    /// Date of holiday
    /// </summary>
    public string Date { get; set; } = string.Empty;
    /// <summary>
    /// Country code of the holiday
    /// </summary>
    public string CountryCode { get; set; } = string.Empty;
    #endregion properties
}