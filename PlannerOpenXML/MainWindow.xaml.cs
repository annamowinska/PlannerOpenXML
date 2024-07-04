using Microsoft.Extensions.DependencyInjection;
using Notification.Wpf;
using PlannerOpenXML.Converters;
using PlannerOpenXML.Services;
using PlannerOpenXML.ViewModel;

namespace PlannerOpenXML;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow
{
    #region constructors
    public MainWindow()
    {
        ServiceContainer.Services = ConfigureInternalServices();
        DataContext = ServiceContainer.Services.GetRequiredService<MainViewModel>();
        InitializeComponent();
    }
    #endregion constructors

    #region private methods
    private static ServiceProvider ConfigureInternalServices()
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
        services.AddTransient<DialogService>();
        services.AddTransient<MainViewModel>();

        return services.BuildServiceProvider();
    }
    #endregion private methods
}