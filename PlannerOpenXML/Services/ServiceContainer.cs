using Microsoft.Extensions.DependencyInjection;

namespace PlannerOpenXML.Services;

public static class ServiceContainer
{
    #region properties
    public static IServiceProvider? Services { get; set; }
    #endregion properties

    #region methods
    public static TService GetService<TService>() =>
        (TService)GetService(typeof(TService));

    public static object GetService(Type serviceInterface)
    {
        return Services is null
            ? throw new ApplicationException("Services not initialized properly.")
            : Services.GetRequiredService(serviceInterface);
    }
    #endregion methods
}