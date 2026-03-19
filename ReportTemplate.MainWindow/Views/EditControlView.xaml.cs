using System;
using System.Windows.Controls;
using Caliburn.Micro;
using ReportTemplate.MainWindow.Events;
using ReportTemplate.MainWindow.ViewModels;

namespace ReportTemplate.MainWindow.Views;

public partial class EditControlView : UserControl,
    IHandle<RequestWordPanelHandleEvent>,
    IHandle<SubscribePanelResizeEvent>,
    IHandle<UnsubscribePanelResizeEvent>
{
    private readonly IEventAggregator _eventAggregator;
    private System.Windows.Forms.Panel? _panel;
    private EventHandler? _resizeEventHandler;

    public EditControlView(IEventAggregator eventAggregator)
    {
        InitializeComponent();
        _eventAggregator = eventAggregator;
        _eventAggregator.SubscribeOnUIThread(this);

        // 获取 Panel 引用
        _panel = WordHostControl?.Child as System.Windows.Forms.Panel;
    }

    /// <summary>
    /// 处理请求 Word 嵌入面板句柄事件
    /// </summary>
    public System.Threading.Tasks.Task HandleAsync(RequestWordPanelHandleEvent message, System.Threading.CancellationToken cancellationToken)
    {
        if (_panel != null)
        {
            message.PanelHandle = _panel.Handle;
        }
        return System.Threading.Tasks.Task.CompletedTask;
    }

    /// <summary>
    /// 处理订阅 Panel Resize 事件
    /// </summary>
    public System.Threading.Tasks.Task HandleAsync(SubscribePanelResizeEvent message, System.Threading.CancellationToken cancellationToken)
    {
        if (_panel != null)
        {
            // 先取消之前的订阅（如果有）
            if (_resizeEventHandler != null)
            {
                _panel.Resize -= _resizeEventHandler;
            }

            // 创建并保存新的事件处理器
            _resizeEventHandler = (s, e) => Panel_Resize(message.WordWnd, message.PanelHandle);
            _panel.Resize += _resizeEventHandler;
        }
        return System.Threading.Tasks.Task.CompletedTask;
    }

    /// <summary>
    /// 处理取消订阅 Panel Resize 事件
    /// </summary>
    public System.Threading.Tasks.Task HandleAsync(UnsubscribePanelResizeEvent message, System.Threading.CancellationToken cancellationToken)
    {
        if (_panel != null && _resizeEventHandler != null)
        {
            _panel.Resize -= _resizeEventHandler;
            _resizeEventHandler = null;
        }
        return System.Threading.Tasks.Task.CompletedTask;
    }

    /// <summary>
    /// Panel 大小变化时同步调整 Word 窗口
    /// </summary>
    private void Panel_Resize(IntPtr wordWnd, IntPtr panelHandle)
    {
        if (wordWnd != IntPtr.Zero && panelHandle != IntPtr.Zero)
        {
            EditControlViewModel.NativeMethods.GetClientRect(panelHandle, out var rect);
            EditControlViewModel.NativeMethods.MoveWindow(wordWnd, 0, 0, rect.Right, rect.Bottom, true);
        }
    }
}
