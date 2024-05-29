using PlannerOpenXML.Model;

namespace PlannerOpenXML.Services;

public class AddedMilestoneService
{
    #region fields
    private List<AddedMilestone> m_AddedMilestones = new List<AddedMilestone>();
    #endregion fields

    #region methods
    public void AddMilestone(AddedMilestone milestone)
    {
        m_AddedMilestones.Add(milestone);
    }

    public List<AddedMilestone> GetAddedMilestones()
    {
        return m_AddedMilestones;
    }

    public void ClearAddedMilestones()
    {
        m_AddedMilestones.Clear();
    }
    #endregion methods
}