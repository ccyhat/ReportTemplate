using System;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows;
using Caliburn.Micro;
using MessageBox = System.Windows.MessageBox;
using MessageBoxImage = System.Windows.MessageBoxImage;
using MessageBoxButton = System.Windows.MessageBoxButton;
using ReportTemplate.MainWindow.Events;

namespace ReportTemplate.MainWindow.ViewModels;

public class EditControlViewModel : Screen
{
    private readonly IEventAggregator _eventAggregator;
    private dynamic? _wordApp;
    private dynamic? _wordDoc;
    private string _filePath = string.Empty;
    private string? _selectedItem;

    public string FilePath
    {
        get => _filePath;
        set => Set(ref _filePath, value);
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

  

    public EditControlViewModel(IEventAggregator eventAggregator)
    {
        _eventAggregator = eventAggregator;
        DisplayName = "Word 编辑器";
    }

   

    /// <summary>
    /// 处理模板列表双击事件
    /// </summary>
    public void TemplateList_DoubleClick(string? item)
    {
        if (_wordDoc == null)
        {
            MessageBox.Show("请先打开 Word 文档", "提示",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        InsertBookmarkAtSelection("hello");
    }

    /// <summary>
    /// 在 Word 当前光标位置插入书签
    /// </summary>
    /// <param name="bookmarkName">书签名称</param>
    private void InsertBookmarkAtSelection(string bookmarkName)
    {
        try
        {
            if (_wordDoc == null)
            {
                MessageBox.Show("文档未打开", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 获取当前选区
            var selection = _wordApp.Selection;

            // 如果存在同名书签，先删除
            if (_wordDoc.Bookmarks.Exists(bookmarkName))
            {
                _wordDoc.Bookmarks[bookmarkName].Delete();
            }

            // 在当前选区位置添加书签
            _wordDoc.Bookmarks.Add(bookmarkName, selection);

            MessageBox.Show($"书签 '{bookmarkName}' 已插入到光标位置", "提示",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"插入书签失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
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

            _wordDoc = _wordApp?.Documents.Open(fileName, missing, readOnly,
                missing, missing, missing, missing, missing, missing,
                missing, missing, isVisible, missing, missing, missing, missing);

            // 请求 Panel 句柄并嵌入 Word 窗口
            RequestPanelAndEmbed();
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

    /// <summary>
    /// 保存 Word 文档
    /// </summary>
    public void SaveDocument()
    {
        try
        {
            if (_wordDoc == null)
            {
                MessageBox.Show("没有打开的文档可保存", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _wordDoc.Save();
            MessageBox.Show("文档已保存", "提示",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存文档失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 请求 Panel 句柄并嵌入 Word 窗口
    /// </summary>
    private async void RequestPanelAndEmbed()
    {
        if (_wordApp == null || _wordDoc == null)
            return;

        // 发布事件请求 Panel 句柄（在 UI 线程上）
        var eventMessage = new RequestWordPanelHandleEvent();
        await _eventAggregator.PublishOnUIThreadAsync(eventMessage);

        if (eventMessage.PanelHandle != IntPtr.Zero)
        {
            EmbedWordIntoWindow(eventMessage.PanelHandle);
        }
        else
        {
            MessageBox.Show("无法获取 Panel 句柄", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }



    /// <summary>
    /// 嵌入 Word 窗口到 Panel
    /// </summary>
    private void EmbedWordIntoWindow(IntPtr panelHandle)
    {
        try
        {
            // 获取 Word 应用程序的主窗口句柄
#pragma warning disable CS8602
            IntPtr wordWnd = (IntPtr)_wordApp.ActiveWindow.Hwnd;
#pragma warning restore CS8602

            // 将 Word 窗口设置为 Panel 的子窗口
            SetWordParent(wordWnd, panelHandle);

            _wordApp.Visible = true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"嵌入 Word 失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 设置 Word 窗口的父窗口并调整样式
    /// </summary>
    private void SetWordParent(IntPtr wordWnd, IntPtr parentHandle)
    {
        // 将 Word 窗口设置为 Panel 的子窗口
        NativeMethods.SetParent(wordWnd, parentHandle);

        // 移除 Word 窗口的标题栏和边框样式
        int style = NativeMethods.GetWindowLong(wordWnd, -16);
        style &= ~0x00C00000; // 移除 WS_CAPTION
        style &= ~0x00080000; // 移除 WS_BORDER
        NativeMethods.SetWindowLong(wordWnd, -16, style);

        // 设置 Word 窗口大小以填满 Panel
        NativeMethods.GetClientRect(parentHandle, out var rect);
        NativeMethods.MoveWindow(wordWnd, 0, 0, rect.Right, rect.Bottom, true);

        NativeMethods.ShowWindow(wordWnd, 5); // SW_SHOW

        // 请求 View 订阅 Panel 的 Resize 事件
        var resizeMessage = new SubscribePanelResizeEvent { WordWnd = wordWnd, PanelHandle = parentHandle };
        _eventAggregator.PublishOnUIThreadAsync(resizeMessage);
    }

   
    

    /// <summary>
    /// 关闭 Word 文档
    /// </summary>
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

            // 取消订阅 Panel 的 Resize 事件
            var unsubscribeMessage = new UnsubscribePanelResizeEvent();
            _eventAggregator.PublishOnUIThreadAsync(unsubscribeMessage);

            FilePath = string.Empty;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"关闭 Word 失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 清理 Word 资源
    /// </summary>
    public void Cleanup()
    {
        CloseWord();
    }

    public static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

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

        [DllImport("user32.dll")]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
    }
}
