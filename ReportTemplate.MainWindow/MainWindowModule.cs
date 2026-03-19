using Autofac;
using System.Reflection;

namespace ReportTemplate.MainWindow;

public class MainWindowModule : Autofac.Module
{
    protected override void Load(ContainerBuilder builder)
    {
        // 注册 View
        builder.RegisterAssemblyTypes(Assembly.GetExecutingAssembly())
               .Where(t => t.Name.EndsWith("View"))
               .AsSelf()
               .InstancePerDependency();

        // 单例注册 ViewModel（按依赖顺序：被依赖的先注册）
        builder.RegisterType<ViewModels.EditControlViewModel>()
               .AsSelf()
               .AsImplementedInterfaces()
               .SingleInstance();

        builder.RegisterType<ViewModels.XMLTreeViewModel>()
               .AsSelf()
               .AsImplementedInterfaces()
               .SingleInstance();

        // RibbonControlViewModel 依赖 EditControlViewModel，所以后注册
        builder.RegisterType<ViewModels.RibbonControlViewModel>()
               .AsSelf()
               .AsImplementedInterfaces()
               .SingleInstance();

        // MainWindowViewModel 依赖以上所有 ViewModel，最后注册
        builder.RegisterType<ViewModels.MainWindowViewModel>()
               .AsSelf()
               .AsImplementedInterfaces()
               .SingleInstance();
    }
}
