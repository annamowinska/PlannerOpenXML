using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace PlannerOpenXML.Converters;

public class SecondCountryLabelVisibilityConverter : IValueConverter
{
    #region properties
    public Visibility LabelCollapsed { get; set; } = Visibility.Collapsed;
    public Visibility LabelVisible { get; set; } = Visibility.Visible;
    public Visibility Default { get; set; } = Visibility.Collapsed;
    #endregion properties

    #region methods
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not bool boolean)
            return Default;

        return boolean ? LabelVisible : LabelCollapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
    #endregion methods
}