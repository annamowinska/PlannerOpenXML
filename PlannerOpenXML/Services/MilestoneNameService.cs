using PlannerOpenXML.Model;

namespace PlannerOpenXML.Services;

public class MilestoneNameService
{
    #region methods
    public string GetMilestoneName(DateOnly date, IEnumerable<Milestone> milestones)
    {
        foreach (var milestone in milestones)
        {
            if (milestone.MilestoneDate == date)
            {
                return milestone.MilestoneText;
            }
        }
        return string.Empty;
    }
    #endregion methods
}
