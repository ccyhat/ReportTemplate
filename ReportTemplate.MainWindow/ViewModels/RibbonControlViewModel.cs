using Caliburn.Micro;
using Microsoft.Win32;
using System.Windows;

namespace ReportTemplate.MainWindow.ViewModels;

public class RibbonControlViewModel : Screen
{
    private readonly EditControlViewModel _editControlViewModel;

    public RibbonControlViewModel(EditControlViewModel editControlViewModel)
    {
        _editControlViewModel = editControlViewModel;
        DisplayName = "RibbonControl";
    }

    public void OpenFileClicked(object source, object eventArgs)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Word 文档 (*.docx;*.doc)|*.docx;*.doc|所有文件 (*.*)|*.*",
            Title = "打开 Word 文档"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            _editControlViewModel.OpenDocument(openFileDialog.FileName);
        }
    }

    public void CloseFileClicked(object source, object eventArgs)
    {
        _editControlViewModel.CloseDocument();
    }

    public void SaveFileClicked(object source, object eventArgs)
    {
        MessageBox.Show("保存文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    public void SaveFileAsClicked(object source, object eventArgs)
    {
        MessageBox.Show("另存为", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
