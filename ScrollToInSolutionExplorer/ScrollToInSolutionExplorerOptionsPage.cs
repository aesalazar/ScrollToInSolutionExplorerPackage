// ReSharper disable once RedundantNullableDirective - Necessary for VS Extensions

#nullable enable

using System.ComponentModel;
using System.Windows;
using Microsoft.VisualStudio.Shell;

namespace ScrollToInSolutionExplorer;

/// <summary>
/// Dialog to preset in the Visual Studio Options pages.
/// </summary>
internal sealed class ScrollToInSolutionExplorerOptionsPage : UIElementDialogPage
{
    public const string IsVisibleInToolMenuHeader = "Show in Visual Studio Tools Menu";
    public const string IsVisibleInDocumentTabMenuHeader = "Show in Document Tab Context Menu";
    
    private ScrollToInSolutionExplorerCommandManager? _commandManager;
    private ScrollToInSolutionExplorerOptionsControl? _optionsControl;

    protected override UIElement Child => _optionsControl ??= new ScrollToInSolutionExplorerOptionsControl();

    /// <summary>
    /// Indicates if the icon should be displayed in the Visual Studio Tool menu.
    /// </summary>
    [Category("Visibility")]
    [DisplayName(IsVisibleInToolMenuHeader)]
    [Browsable(true)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
    public bool IsVisibleInToolMenu { get; set; } = true;

    /// <summary>
    /// Indicates if the icon should be displayed in the Document Tab context menu.
    /// </summary>
    [Category("Visibility")]
    [DisplayName(IsVisibleInDocumentTabMenuHeader)]
    [Browsable(true)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
    public bool IsVisibleInDocumentTabMenu { get; set; } = true;

    /// <summary>
    /// Sets up for user invoked changes.
    /// </summary>
    /// <param name="commandManager">Command Manager to call to update visibilities.</param>
    public void Initialize(ScrollToInSolutionExplorerCommandManager commandManager)
    {
        _commandManager = commandManager;
        _commandManager.ApplyUserSettings(this);
    }

    protected override void OnActivate(CancelEventArgs e)
    {
        base.OnActivate(e);
        LoadSettingsFromStorage();

        if (_optionsControl is not null)
        {
            _optionsControl.IsToolMenuVisible = IsVisibleInToolMenu;
            _optionsControl.IsDocumentTabMenuVisible = IsVisibleInDocumentTabMenu;
        }
    }

    protected override void OnApply(PageApplyEventArgs e)
    {
        base.OnApply(e);

        if (_optionsControl is not null)
        {
            IsVisibleInToolMenu = _optionsControl.IsToolMenuVisible;
            IsVisibleInDocumentTabMenu = _optionsControl.IsDocumentTabMenuVisible;
        }

        _commandManager?.ApplyUserSettings(this);
        SaveSettingsToStorage();
    }
}