using Caliburn.Micro;
using System.Collections.ObjectModel;
using Screen = Caliburn.Micro.Screen;

namespace ReportTemplate.MainWindow.ViewModels;

public class XmlTreeNodeViewModel
{
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public ObservableCollection<XmlTreeNodeViewModel> Children { get; set; } = new();
}

public class XMLTreeViewModel : Screen
{
    public ObservableCollection<XmlTreeNodeViewModel> Nodes { get; } = new();

    public XMLTreeViewModel()
    {
        DisplayName = "XML 结构";
        InitializeSampleData();
    }

    private void InitializeSampleData()
    {
        Nodes.Add(new XmlTreeNodeViewModel { Name = "根节点", Type = "Root" });
        Nodes[0].Children.Add(new XmlTreeNodeViewModel { Name = "子节点 1", Type = "Element" });
        Nodes[0].Children.Add(new XmlTreeNodeViewModel { Name = "子节点 2", Type = "Element" });
        Nodes[0].Children[0].Children.Add(new XmlTreeNodeViewModel { Name = "孙节点 1-1", Type = "Attribute" });
    }
}
