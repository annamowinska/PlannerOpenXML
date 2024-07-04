using System.Windows;

namespace PlannerOpenXML.AttachedProperties;

public class MarginProperties
{
    #region dependency properties
    public static readonly DependencyProperty RightProperty = DependencyProperty.RegisterAttached(
        "Right",
        typeof(double),
        typeof(MarginProperties),
        new UIPropertyMetadata(OnRightPropertyChanged));
    #endregion dependency properties

    #region methods
    public static double GetRight(FrameworkElement element)
    {
        return (double)element.GetValue(RightProperty);
    }

    public static void SetRight(FrameworkElement element, double value)
    {
        element.SetValue(RightProperty, value);
    }
    #endregion methods

    #region private methods
    private static void OnRightPropertyChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
    {
        if (obj is FrameworkElement element && args.NewValue is double value)
        {
            var margin = element.Margin;
            margin.Right = value;
            element.Margin = margin;
        }
    }
    #endregion private methods
}
