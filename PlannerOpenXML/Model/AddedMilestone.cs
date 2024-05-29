namespace PlannerOpenXML.Model;

public class AddedMilestone
{
    #region properties
    public string AddedMilestoneText { get; set; } = string.Empty;
    public DateOnly AddedMilestoneDate { get; set; } = new DateOnly();
    #endregion properties
}