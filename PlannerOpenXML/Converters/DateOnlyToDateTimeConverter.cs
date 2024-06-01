using System.Globalization;
using System.Windows.Data;

namespace PlannerOpenXML.Converters;

public class DateOnlyToDateTimeConverter : IValueConverter
{
    #region methods
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not DateOnly dateOnly)
            return DateTime.Now;

        return dateOnly.ToDateTime(TimeOnly.MinValue);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not DateTime dateTime)
            return DateOnly.FromDateTime(DateTime.Now);

        return DateOnly.FromDateTime(dateTime);
    }
    #endregion methods
}
