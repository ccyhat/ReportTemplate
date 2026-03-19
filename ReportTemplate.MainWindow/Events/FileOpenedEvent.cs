namespace ReportTemplate.MainWindow.Events;

/// <summary>
/// 文件打开事件，用于在 ViewModel 间传递文件内容
/// </summary>
public class FileOpenedEvent
{
    public string Content { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
}
