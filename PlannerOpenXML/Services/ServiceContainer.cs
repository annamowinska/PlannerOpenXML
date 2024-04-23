namespace PlannerOpenXML.Services;

public static class ServiceContainer
{
    public static IServiceProvider Services { get; set; }

    public static TService GetService<TService>() =>
        (TService)GetService(typeof(TService));

    public static object GetService(Type serviceInterface)
    {
        return Services.GetService(serviceInterface);
    }
}

