using Microsoft.Extensions.DependencyInjection;
using Notification.Wpf;
using PlannerOpenXML.Converters;
using PlannerOpenXML.Model;
using PlannerOpenXML.Services;
using PlannerOpenXML.ViewModel;
using System.Windows;

namespace PlannerOpenXML;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        ServiceContainer.Services = ConfigureInternalServices();
        DataContext = ServiceContainer.Services.GetRequiredService<MainViewModel>();
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
    }

    #region private methods
    private IServiceProvider ConfigureInternalServices()
    {
        var services = new ServiceCollection();

        services.AddSingleton<ICountryListService, CountryListService>();
        services.AddTransient<IHolidayConverter, NagerHolidayConverter>();
        services.AddTransient<IApiService, ApiNagerService>();
        services.AddTransient<HolidayNameService>();
        services.AddTransient<HolidayCacheService>();
        services.AddTransient<INotificationManager, NotificationManager>();
        services.AddSingleton<NotificationService>();
        services.AddTransient<PlannerStyleService>();
        services.AddTransient<PlannerGenerator>();
        services.AddTransient<DialogService>();
        services.AddTransient<MainViewModel>();

        return services.BuildServiceProvider();
    }
    #endregion private methods
}