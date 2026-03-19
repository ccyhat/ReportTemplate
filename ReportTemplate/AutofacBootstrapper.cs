using Autofac;
using Caliburn.Micro;
using ReportTemplate.MainWindow;
using ReportTemplate.MainWindow.ViewModels;
using ReportTemplate.MainWindow.Views;
using System.Reflection;

namespace ReportTemplate;

public class AutofacBootstrapper : BootstrapperBase
{
    private IContainer _container = null!;

    public AutofacBootstrapper()
    {
        Initialize();
    }

    protected override async void OnStartup(object sender, System.Windows.StartupEventArgs e)
    {
        await DisplayRootViewForAsync<MainWindowViewModel>();
    }

    protected override void Configure()
    {
        var builder = new ContainerBuilder();

        // 1. 注册 Caliburn.Micro 核心服务（必须先注册）
        builder.RegisterType<WindowManager>().As<IWindowManager>().SingleInstance();
        builder.RegisterType<EventAggregator>().As<IEventAggregator>().SingleInstance();

        // 2. 注册 MainWindow 模块
        builder.RegisterModule<MainWindowModule>();

        // 3. 扫描并添加引用的程序集
        ScanAndAddReferencedAssemblies();

        // 4. 构建容器
        _container = builder.Build();
    }

    // 重写 GetInstance 方法，从 Autofac 容器中解析实例
    protected override object GetInstance(Type service, string key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            if (_container.TryResolve(service, out object? instance))
            {
                return instance!;
            }
            throw new Exception($"Could not resolve type: {service.FullName}");
        }

        if (_container.TryResolveNamed(key, service, out object? namedInstance))
        {
            return namedInstance!;
        }
        throw new Exception($"Could not resolve type: {service.FullName} with key: {key}");
    }

    // 重写 GetAllInstances 方法
    protected override IEnumerable<object> GetAllInstances(Type service)
    {
        var type = typeof(IEnumerable<>).MakeGenericType(service);
        return (IEnumerable<object>)_container.Resolve(type);
    }

    // Caliburn 构建实例时，注入依赖
    protected override void BuildUp(object instance)
    {
        _container.InjectProperties(instance);
    }

    private void ScanAndAddReferencedAssemblies()
    {
        var mainAssembly = Assembly.GetExecutingAssembly();
        var referencedAssemblies = mainAssembly.GetReferencedAssemblies();

        Console.WriteLine($"主程序集：{mainAssembly.FullName}");
        Console.WriteLine($"引用的程序集数量：{referencedAssemblies.Length}");

        foreach (var assemblyName in referencedAssemblies)
        {
            try
            {
                var assembly = Assembly.Load(assemblyName);
                Console.WriteLine($"加载程序集：{assembly.FullName}");
                AssemblySource.Instance.Add(assembly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"无法加载程序集：{assemblyName.Name}, 错误：{ex.Message}");
            }
        }
    }
}
