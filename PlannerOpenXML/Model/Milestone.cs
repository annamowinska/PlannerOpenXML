namespace PlannerOpenXML.Model;

public class Milestone
{
    #region properties
    public string MilestoneText { get; set; } = string.Empty;
    public DateOnly MilestoneDate { get; set; } = new DateOnly();
    #endregion properties
}
