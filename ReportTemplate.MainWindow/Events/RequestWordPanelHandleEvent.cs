namespace ReportTemplate.MainWindow.Events;

/// <summary>
/// 请求 Word 嵌入面板句柄事件，用于 View 向 ViewModel 提供 Panel 句柄
/// </summary>
public class RequestWordPanelHandleEvent
{
    public System.IntPtr PanelHandle { get; set; }
}
