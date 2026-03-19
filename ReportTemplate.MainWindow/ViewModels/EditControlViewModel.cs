using System;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms.Integration;
using Caliburn.Micro;
using Panel = System.Windows.Forms.Panel;
using MessageBox = System.Windows.MessageBox;
using MessageBoxImage = System.Windows.MessageBoxImage;
using MessageBoxButton = System.Windows.MessageBoxButton;

namespace ReportTemplate.MainWindow.ViewModels;

public class EditControlViewModel : Screen
{
    private dynamic? _wordApp;
    private dynamic? _wordDoc;
    private WindowsFormsHost? _host;
    private Panel? _panel;
    private string _filePath = string.Empty;
    private string? _selectedItem;

    public string FilePath
    {
        get => _filePath;
        set
        {
            _filePath = value;
            NotifyOfPropertyChange(() => FilePath);
        }
    }

    /// <summary>
    /// 左侧模板列表数据源
    /// </summary>
    public ObservableCollection<string> TemplateList { get; } = new()
    {
        "模板 1 - 报告封面",
        "模板 2 - 目录页",
        "模板 3 - 正文内容",
        "模板 4 - 数据表格",
        "模板 5 - 附录"
    };

    /// <summary>
    /// 选中的列表项
    /// </summary>
    public string? SelectedItem
    {
        get => _selectedItem;
        set
        {
            _selectedItem = value;
            NotifyOfPropertyChange(() => SelectedItem);
            OnSelectedItemChanged();
        }
    }

    public EditControlViewModel()
    {
        DisplayName = "Word 编辑器";
    }

    /// <summary>
    /// 处理列表项选中事件
    /// </summary>
    private void OnSelectedItemChanged()
    {
        if (!string.IsNullOrEmpty(SelectedItem))
        {
            // 暂时仅显示提示
            // MessageBox.Show($"已选择：{SelectedItem}", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    /// <summary>
    /// 打开 Word 文档并嵌入到编辑区域
    /// </summary>
    public void OpenDocument(string filePath)
    {
        try
        {
            // 关闭已打开的 Word 文档
            CloseWord();

            FilePath = filePath;

            // 通过 COM 创建 Word 应用程序实例
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                MessageBox.Show("未找到 Microsoft Word，请确保已安装 Word。", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _wordApp = Activator.CreateInstance(wordType);
            

            // 打开文档
            object fileName = filePath;
            object readOnly = false;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;

            _wordDoc = _wordApp.Documents.Open(fileName, missing, readOnly,
                missing, missing, missing, missing, missing, missing,
                missing, missing, isVisible, missing, missing, missing, missing);

            // 获取 Word 窗口句柄并嵌入
            EmbedWordIntoWindow();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"打开 Word 文档失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            CloseWord();
            throw;
        }
    }

    /// <summary>
    /// 关闭 Word 文档
    /// </summary>
    public void CloseDocument()
    {
        CloseWord();
    }

    private void EmbedWordIntoWindow()
    {
        if (_wordApp == null || _wordDoc == null)
            return;

        try
        {
            // 获取 Word 应用程序的主窗口句柄
            IntPtr wordWnd = (IntPtr)_wordApp.ActiveWindow.Hwnd;

            // 创建 WindowsFormsHost 来承载 Word
            _host = new WindowsFormsHost();
            _panel = new Panel
            {
                Dock = System.Windows.Forms.DockStyle.Fill
            };

            // 将 Word 窗口设置为 Panel 的子窗口
            NativeMethods.SetParent(wordWnd, _panel.Handle);

            // 移除 Word 窗口的标题栏和边框样式
            int style = NativeMethods.GetWindowLong(wordWnd, -16);
            style &= ~0x00C00000; // 移除 WS_CAPTION
            style &= ~0x00080000; // 移除 WS_BORDER
            NativeMethods.SetWindowLong(wordWnd, -16, style);

            // 设置 Word 窗口大小以填满 Panel
            NativeMethods.MoveWindow(wordWnd, 0, 0, _panel.Width, _panel.Height, true);

            // 监听 Panel 大小变化，同步调整 Word 窗口
            _panel.Resize += (s, e) =>
            {
                if (_wordApp?.ActiveWindow != null)
                {
                    IntPtr wnd = (IntPtr)_wordApp.ActiveWindow.Hwnd;
                    NativeMethods.MoveWindow(wnd, 0, 0, _panel.Width, _panel.Height, true);
                }
            };

            NativeMethods.ShowWindow(wordWnd, 5); // SW_SHOW

            _host.Child = _panel;
            _wordApp.Visible = true;
            // 通知 View 更新内容
            NotifyOfPropertyChange(() => WordHost);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"嵌入 Word 失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CloseWord()
    {
        try
        {
            if (_wordDoc != null)
            {
                object saveChanges = false;
                _wordDoc.Close(saveChanges);
                _wordDoc = null;
            }

            if (_wordApp != null)
            {
                _wordApp.Quit();
                Marshal.ReleaseComObject(_wordApp);
                _wordApp = null;
            }

            if (_host != null)
            {
                _host.Child = null;
                _host = null;
            }

            if (_panel != null)
            {
                _panel.Dispose();
                _panel = null;
            }

            FilePath = string.Empty;
            NotifyOfPropertyChange(() => WordHost);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"关闭 Word 失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 获取 Word 嵌入容器
    /// </summary>
    public object? WordHost => _host;

    protected override void OnViewLoaded(object view)
    {
        base.OnViewLoaded(view);
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    }
}
