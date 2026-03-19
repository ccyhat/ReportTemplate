namespace ReportTemplate.MainWindow.Events;

/// <summary>
/// 订阅 Panel Resize 事件，用于 ViewModel 通知 View 订阅 Panel 大小变化
/// </summary>
public class SubscribePanelResizeEvent
{
    public System.IntPtr WordWnd { get; set; }
    public System.IntPtr PanelHandle { get; set; }
}
