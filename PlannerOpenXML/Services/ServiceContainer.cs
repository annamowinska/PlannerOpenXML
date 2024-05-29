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
        return Services.GetService(serviceInterface);
    }
    #endregion methods
}