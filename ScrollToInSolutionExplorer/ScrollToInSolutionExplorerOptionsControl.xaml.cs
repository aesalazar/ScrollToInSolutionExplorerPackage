namespace ScrollToInSolutionExplorer;

public partial class ScrollToInSolutionExplorerOptionsControl
{
    public ScrollToInSolutionExplorerOptionsControl()
    {
        InitializeComponent();
    }

    public bool IsToolMenuVisible
    {
        get => IsVisibleInToolMenu.IsChecked ?? true;
        set => IsVisibleInToolMenu.IsChecked = value;
    }

    public bool IsDocumentTabMenuVisible
    {
        get => IsVisibleInDocumentTabMenu.IsChecked ?? true;
        set => IsVisibleInDocumentTabMenu.IsChecked = value;
    }
}