using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace PlannerOpenXML.Converters;

public class NullToVisibilityConverter : IValueConverter
{
    #region properties
    public Visibility WhenTrue { get; set; } = Visibility.Visible;
    public Visibility WhenFalse { get; set; } = Visibility.Collapsed;
    #endregion properties

    #region methods
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is null ? WhenTrue : WhenFalse;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
    #endregion methods
}
