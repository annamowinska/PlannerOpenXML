namespace PlannerOpenXML.Model;

public class Holiday
{
    #region properties
    /// <summary>
    /// Name of the holiday
    /// </summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>
    /// Date of holiday
    /// </summary>
    public DateOnly Date { get; set; } = new DateOnly();
    #endregion properties
}
