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

    public void ShowNotificationIsSameCountriesSelected()
    {
        var content = new NotificationContent()
        {
            Title = "Error",
            Message = "The same country was chosen.",
            Type = NotificationType.Warning,
            TrimType = NotificationTextTrimType.NoTrim,
            CloseOnClick = true,
            Background = new SolidColorBrush(Colors.Red),
            Foreground = new SolidColorBrush(Colors.White),
        };
        m_NotificationManager.Show(content);
    }

    public void ShowNotificationCountryInput()
    {
        var content = new NotificationContent()
        {
            Title = "Error",
            Message = "The entered country is not on the list. Try selecting a country from the drop-down list.",
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