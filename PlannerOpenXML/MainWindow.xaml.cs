using Microsoft.Extensions.DependencyInjection;
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
        ServiceContainer.Services = ConfigureInternalServices();
        InitializeComponent();
        DataContext = ConfigureInternalServices().GetRequiredService<MainViewModel>();
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
    }

    private IServiceProvider ConfigureInternalServices()
    {
        var services = new ServiceCollection();

        services.AddTransient<IHolidayConverter, NagerHolidayConverter>();
        services.AddTransient<IApiService, ApiNagerService>();
        services.AddTransient<HolidayNameService>();
        services.AddTransient<HolidayCacheService>();
        services.AddTransient<PlannerStyleService>();
        services.AddTransient<PlannerGenerator>();
        services.AddTransient<MainViewModel>();

        return services.BuildServiceProvider();
    }
}