using PlannerOpenXML.Model;
using System.IO;
using System.Text.Json;

namespace PlannerOpenXML.Services
{
    public class MilestoneListReader
    {
        private readonly string m_MilestonesFilePath;

        public MilestoneListReader(string milestonesFilePath = "../../../Resources/milestones.json")
        {
            m_MilestonesFilePath = milestonesFilePath;
        }

        public List<Milestone> LoadMilestones()
        {
            if (!File.Exists(m_MilestonesFilePath))
            {
                return new List<Milestone>();
            }
                
            var json = File.ReadAllText(m_MilestonesFilePath);
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                var milestones = JsonSerializer.Deserialize<List<Milestone>>(json, options);

                if (milestones == null)
                {
                    return new List<Milestone>();
                }

                return milestones;
            }
            catch (JsonException ex)
            {
                return new List<Milestone>();
            }
        }
    }
}