using Notification.Wpf;
using System.Windows.Media;

namespace PlannerOpenXML.Services;

public class NotificationService
{
    #region fields
    private readonly INotificationManager m_NotificationManager;
    #endregion fields

    #region constructor
    public NotificationService(INotificationManager notificationManager)
    {
        m_NotificationManager = notificationManager;
    }
    #endregion constructor

    #region methods
    public void ShowNotificationErrorYearAndFirstMonthInput()
    {
        var content = new NotificationContent()
        {
            Title = "Error",
            Message = "Please enter only numbers.",
            Type = NotificationType.Warning,
            TrimType = NotificationTextTrimType.NoTrim,
            CloseOnClick = true,
            Background = new SolidColorBrush(Colors.Red),
            Foreground = new SolidColorBrush(Colors.White),
        };
        m_NotificationManager.Show(content);
    }

    public void ShowNotificationAddedMilestone(string milestoneText, string milestoneDate)
    {
        var content = new NotificationContent()
        {
            Title = "Success",
            Message = $"Milestone '{milestoneText}' added on {milestoneDate}.",
            Type = NotificationType.Information,
            TrimType = NotificationTextTrimType.NoTrim,
            CloseOnClick = true,
            Background = new SolidColorBrush(Colors.Green),
            Foreground = new SolidColorBrush(Colors.White),
        };
        m_NotificationManager.Show(content);
    }

    public void ShowNotificationErrorMilestoneDateInput()
    {
        var content = new NotificationContent()
        {
            Title = "Error",
            Message = "Please enter date in format 'DD-MM-YYYY'.",
            Type = NotificationType.Warning,
            TrimType = NotificationTextTrimType.NoTrim,
            CloseOnClick = true,
            Background = new SolidColorBrush(Colors.Red),
            Foreground = new SolidColorBrush(Colors.White),
        };
        m_NotificationManager.Show(content);
    }

    public void ShowNotificationErrorMilestoneAndMilestoneDateInput()
    {
        var content = new NotificationContent()
        {
            Title = "Error",
            Message = "Please select milestone and enter date in format 'DD-MM-YYYY'.",
            Type = NotificationType.Warning,
            TrimType = NotificationTextTrimType.NoTrim,
            CloseOnClick = true,
            Background = new SolidColorBrush(Colors.Red),
            Foreground = new SolidColorBrush(Colors.White),
        };
        m_NotificationManager.Show(content);
    }
    #endregion methods
}