using PlannerOpenXML.Model;
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace PlannerOpenXML.Services;

public class MilestoneService
{
    #region fields
    private readonly string m_FilePath;
    #endregion fields

    #region constructors
    public MilestoneService()
    {
        var projectDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var resourcesDirectory = Path.Combine(projectDirectory, "..", "..", "..", "Resources");
        m_FilePath = Path.Combine(resourcesDirectory, "milestones.json");

        if (!Directory.Exists(resourcesDirectory))
        {
            Directory.CreateDirectory(resourcesDirectory);
        }
    }
    #endregion constructors

    #region methods
    public void SaveMilestonesToFile(List<Milestone> milestones)
    {
        var json = JsonSerializer.Serialize(milestones, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(m_FilePath, json);
    }
    #endregion methods
}
