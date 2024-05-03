using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace PlannerOpenXML.Services;
public class InternetAvailabilityService
{
    #region methods
    public static bool IsInternetAvailable()
    {
        try
        {
            Ping ping = new Ping();
            PingReply reply = ping.Send("8.8.8.8");
            return reply.Status == IPStatus.Success;
        }
        catch (PingException)
        {
            return false;
        }
    }
    #endregion methods
}

