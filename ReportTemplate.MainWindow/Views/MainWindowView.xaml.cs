using System.ComponentModel;
using System.Windows;
using ReportTemplate.MainWindow.ViewModels;

namespace ReportTemplate.MainWindow.Views;

public partial class MainWindowView : Window
{
    public MainWindowView()
    {
        InitializeComponent();
        Closing += MainWindowView_Closing;
    }

    private void MainWindowView_Closing(object? sender, CancelEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.Cleanup();
        }
    }
}
