# ReportTemplate 项目上下文

## 项目概述

**ReportTemplate** 是一个基于 **.NET 8.0** 和 **WPF (Windows Presentation Foundation)** 的 Windows 桌面应用程序模板项目。项目采用 **Caliburn.Micro** 作为 MVVM 框架，使用 **Autofac** 作为依赖注入容器，实现了模块化的架构设计。

### 技术栈

| 技术 | 版本/说明 |
|------|-----------|
| **框架** | .NET 8.0 Windows |
| **UI 框架** | WPF |
| **语言** | C# |
| **MVVM 框架** | Caliburn.Micro 5.0.258 |
| **DI 容器** | Autofac 9.1.0 |
| **Nullable 参考类型** | 已启用 |
| **隐式 Using** | 主项目已启用，类库项目已禁用 |

### 项目结构

```
ReportTemplate/
├── ReportTemplate.sln              # Visual Studio 解决方案
├── QWEN.md                         # 项目上下文文档
├── .gitignore                      # Git 忽略配置
├── ReportTemplate/                 # 主应用程序项目 (启动项目)
│   ├── App.xaml(.cs)              # 应用程序入口
│   ├── AutofacBootstrapper.cs     # Caliburn.Micro + Autofac 启动引导器
│   ├── AssemblyInfo.cs            # 程序集信息
│   └── ReportTemplate.csproj      # 项目文件
└── ReportTemplate.MainWindow/      # 类库项目 (UI 视图模块)
    ├── MainWindowModule.cs        # Autofac 模块，负责注册 View/ViewModel
    ├── ViewModels/
    │   ├── MainWindowViewModel.cs # 主窗口 ViewModel
    │   ├── RibbonControlViewModel.cs  # 工具栏 ViewModel
    │   ├── XMLTreeViewModel.cs    # XML 树形结构 ViewModel
    │   └── EditControlViewModel.cs # Word 编辑器 ViewModel
    ├── Views/
    │   ├── MainWindowView.xaml    # 主窗口视图
    │   ├── RibbonControlView.xaml # 工具栏视图
    │   ├── XMLTreeView.xaml       # XML 树形结构视图
    │   └── EditControlView.xaml   # Word 编辑器视图
    └── Events/
        └── FileOpenedEvent.cs     # 文件打开事件
```

### 项目组成

| 项目 | 类型 | 说明 |
|------|------|------|
| `ReportTemplate` | WinExe | 主应用程序入口，包含启动引导和 DI 配置 |
| `ReportTemplate.MainWindow` | Library | 独立的 UI 视图类库，包含 Views、ViewModels 和 Events |

## 架构模式

项目采用 **MVVM + 依赖注入** 的架构模式：

1. **Caliburn.Micro** 提供 MVVM 基础设施（Screen 基类、ViewLocator、事件聚合器等）
2. **Autofac** 负责依赖注入和生命周期管理
3. **模块化设计**：将 UI 视图分离到独立的类库项目中，便于复用和团队协作

### View-ViewModel 绑定

通过 Caliburn.Micro 的 `View.Model` 绑定实现：

```xaml
<ContentControl cal:View.Model="{Binding EditControlViewModel}" />
```

### 依赖注入配置

`AutofacBootstrapper.cs` 是核心的启动引导类，负责：

1. 注册 Caliburn.Micro 核心服务（`WindowManager`、`EventAggregator`）
2. 注册 MainWindow 模块（`MainWindowModule`）
3. 扫描并加载引用的程序集
4. 提供 `GetInstance`、`GetAllInstances`、`BuildUp` 方法供 Caliburn.Micro 使用

`MainWindowModule.cs` 负责注册模块内的 View 和 ViewModel：

- View：按依赖注册（InstancePerDependency）
- ViewModel：单例注册（SingleInstance），按依赖顺序注册

### 事件聚合

项目使用 `FileOpenedEvent` 作为 ViewModel 间通信的事件模型，通过 Caliburn.Micro 的 `IEventAggregator` 实现松耦合通信。

## 核心功能

### Word 文档编辑

`EditControlViewModel` 实现了 Word 文档的嵌入编辑功能：

- 通过 COM 自动化创建 Word 应用程序实例
- 使用 `WindowsFormsHost` 将 Word 窗口嵌入到 WPF 界面中
- 支持打开、关闭 Word 文档
- 使用 P/Invoke 调整 Word 窗口的样式和布局

### Ribbon 工具栏

`RibbonControlViewModel` 提供文件操作命令：

- `OpenFileClicked` - 打开文件对话框
- `CloseFileClicked` - 关闭当前文档
- `SaveFileClicked` - 保存文件
- `SaveFileAsClicked` - 另存为

## 构建与运行

### 前置条件

- .NET 8.0 SDK
- Visual Studio 2022 (推荐) 或 VS Code
- Microsoft Word (用于 Word 文档编辑功能)

### CLI 命令

```bash
# 还原依赖
dotnet restore

# 构建解决方案
dotnet build

# 运行应用程序
dotnet run --project ReportTemplate/ReportTemplate.csproj

# 发布应用
dotnet publish -c Release -o ./publish
```

### Visual Studio

1. 打开 `ReportTemplate.sln`
2. 设置 `ReportTemplate` 为启动项目
3. 按 `F5` 运行调试

## 开发规范

### 代码约定

- **Nullable 上下文**: 已启用，需处理可空引用
- **隐式 Using**: 主项目已启用，类库项目已禁用（需显式导入命名空间）
- **XAML 命名**: 使用 `*View.xaml` 后缀命名视图文件
- **代码后置**: XAML 对应的 `.xaml.cs` 文件使用 `partial` 类
- **ViewModel 命名**: 使用 `*ViewModel.cs` 后缀，继承 `Screen` 基类

### 架构模式

项目采用 **MVVM + DI** 模式，遵循以下原则：

1. **View-ViewModel 绑定**: 通过 Caliburn.Micro 的 `View.Model` 绑定
2. **依赖注入**: 所有依赖通过构造函数注入，避免 ServiceLocator 反模式
3. **模块化**: 将功能模块放在独立的类库项目中

### 扩展指南

#### 添加新视图

1. 在 `ReportTemplate.MainWindow/Views/` 下创建 `*View.xaml` 和 `*View.xaml.cs`
2. 在 `ReportTemplate.MainWindow/ViewModels/` 下创建对应的 `*ViewModel.cs`
3. 在 `MainWindowModule.Load()` 中注册 View 和 ViewModel

```csharp
// MainWindowModule.cs
builder.RegisterType<YourViewModel>()
       .AsSelf()
       .AsImplementedInterfaces()
       .SingleInstance();
```

#### 添加新服务

1. 定义服务接口
2. 实现服务类
3. 在 `MainWindowModule.Load()` 或 `AutofacBootstrapper.Configure()` 中注册

```csharp
builder.RegisterType<YourService>().As<IYourService>().SingleInstance();
```

#### 添加事件

1. 在 `Events/` 目录下创建事件类
2. 使用 `IEventAggregator` 发布和订阅事件

```csharp
// 发布事件
_eventAggregator.PublishOnUIThread(new FileOpenedEvent { Content = "...", FilePath = "..." });

// 订阅事件
public class MyViewModel : IHandle<FileOpenedEvent>
{
    public void Handle(FileOpenedEvent message) { ... }
}
```

## 注意事项

- 项目使用 .NET 8.0，仅支持 Windows 平台
- Word 编辑功能需要安装 Microsoft Word
- `bin/` 和 `obj/` 目录为构建输出，不应纳入版本控制
- `.vs/` 目录为 Visual Studio 用户特定设置
- `.csproj.user` 文件为用户特定设置，不应纳入版本控制
