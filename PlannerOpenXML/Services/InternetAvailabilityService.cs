using System.Net.NetworkInformation;

namespace PlannerOpenXML.Services;

public class InternetAvailabilityService
{
    #region methods
    public static bool IsInternetAvailable()
    {
        try
        {
            var ping = new Ping();
            var reply = ping.Send("8.8.8.8");
            return reply.Status == IPStatus.Success;
        }
        catch (PingException)
        {
            return false;
        }
    }
    #endregion methods
}