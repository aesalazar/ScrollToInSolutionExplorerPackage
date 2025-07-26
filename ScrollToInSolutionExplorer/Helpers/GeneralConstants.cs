namespace ScrollToInSolutionExplorer.Helpers;

/// <summary>
/// General consts used to interact with Visual Studio.
/// </summary>
internal static class GeneralConstants
{
    /// <summary>
    /// Name to use for Toolbars and Dialogs.
    /// </summary>
    public const string ProjectName = "Scroll To In Solution Explorer";

    /// <summary>
    /// Title shown in the Options Dialog.
    /// </summary>
    public const string OptionsPageHeader = "General";

    /// <summary>
    /// VS Command to locate and select the active document in Solution Explorer.
    /// </summary>
    public const string SyncWithActiveDocumentCommand = "SolutionExplorer.SyncWithActiveDocument";

    /// <summary>
    /// VS Command to collapse outlining in current active document.
    /// </summary>
    public const string CollapseToDefinitionCommand = "Edit.CollapsetoDefinitions";
}