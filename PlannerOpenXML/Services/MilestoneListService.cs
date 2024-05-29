using Newtonsoft.Json;
using System.IO;

public class MilestoneListService
{
    #region fileds
    private const string filePath = "../../../Resources/milestones.json";
    #endregion fields

    #region methods
    public List<string> LoadMilestonesFromJson()
    {
        if (!File.Exists(filePath))
        {
            return new List<string>();
        }

        var json = File.ReadAllText(filePath);
        return JsonConvert.DeserializeObject<List<string>>(json);
    }
    #endregion methods
}