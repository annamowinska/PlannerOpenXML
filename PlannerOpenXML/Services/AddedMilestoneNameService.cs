using PlannerOpenXML.Model;

namespace PlannerOpenXML.Services;

public class AddedMilestoneNameService
{
    #region methods
    public string GetAddedMilestoneName(DateOnly date, List<AddedMilestone> addedMilestones)
    {
        foreach (var addedMilestone in addedMilestones)
        {
            if (addedMilestone.AddedMilestoneDate == date)
            {
                return addedMilestone.AddedMilestoneText;
            }
        }
        return string.Empty;
    }
    #endregion methods
}