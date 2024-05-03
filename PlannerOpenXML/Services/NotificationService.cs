using Notification.Wpf;
using System.Windows.Media;

namespace PlannerOpenXML.Services;

public class NotificationService
{
    private readonly INotificationManager m_NotificationManager;

    public NotificationService(INotificationManager notificationManager)
    {
        m_NotificationManager = notificationManager;
    }

    public void ShowNotification()
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
}

