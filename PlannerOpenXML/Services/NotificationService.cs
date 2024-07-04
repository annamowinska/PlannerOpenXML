using Notification.Wpf;
using System.Windows.Media;

namespace PlannerOpenXML.Services;

public class NotificationService(INotificationManager notificationManager) : INotificationService
{
    #region fields
    private readonly INotificationManager m_NotificationManager = notificationManager;
    private static readonly SolidColorBrush m_Background = new(Colors.Red);
    private static readonly SolidColorBrush m_Foreground = new(Colors.White);
    #endregion fields

    #region methods
    public void NotifyError(string message)
    {
        var content = new NotificationContent()
        {
            Title = "Error",
            Message = message,
            Type = NotificationType.Warning,
            TrimType = NotificationTextTrimType.NoTrim,
            CloseOnClick = true,
            Background = m_Background,
            Foreground = m_Foreground,
        };
        m_NotificationManager.Show(content);
    }
    #endregion methods
}