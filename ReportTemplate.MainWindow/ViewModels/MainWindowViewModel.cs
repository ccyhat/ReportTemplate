using Caliburn.Micro;

namespace ReportTemplate.MainWindow.ViewModels;

public class MainWindowViewModel : Screen
{
    public RibbonControlViewModel RibbonControlViewModel { get; }
    public XMLTreeViewModel XMLTreeViewModel { get; }
    public EditControlViewModel EditControlViewModel { get; }

    public MainWindowViewModel(
        RibbonControlViewModel ribbonControlViewModel,
        XMLTreeViewModel xmlTreeViewModel,
        EditControlViewModel editControlViewModel)
    {
        DisplayName = "ReportTemplate";
        RibbonControlViewModel = ribbonControlViewModel;
        XMLTreeViewModel = xmlTreeViewModel;
        EditControlViewModel = editControlViewModel;
    }

    /// <summary>
    /// 加载 Word 文档到编辑区域
    /// </summary>
    public void LoadWordDocument(string filePath)
    {
        EditControlViewModel.OpenDocument(filePath);
    }

    /// <summary>
    /// 清理所有子 ViewModel 的资源
    /// </summary>
    public void Cleanup()
    {
        EditControlViewModel.Cleanup();
    }
}
