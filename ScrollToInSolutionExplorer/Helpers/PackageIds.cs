namespace ScrollToInSolutionExplorer.Helpers;

/// <summary>
/// Helper class that encapsulates all CommandIDs uses across VS Package.
/// </summary>
/// <remarks>
/// Names here are based on the XML in the .vsct file.
/// </remarks>
internal static class PackageIds
{
    public const int ScrollToInSolutionExplorerMenuGroup = 0x0100;
    public const int ScrollToInSolutionExplorerMenuCommandId = 0x0101;

    public const int ScrollToInSolutionExplorerDocumentTabGroup = 0x1100;
    public const int ScrollToInSolutionExplorerDocumentTabCommandId = 0x1101;

    public const int ScrollToInSolutionExplorerToolbar = 0x1200;
    public const int ScrollToInSolutionExplorerToolbarGroup = 0x1201;
    public const int ScrollToInSolutionExplorerToolbarCommandId = 0x1202;

    public const int BmpScrollTo = 0x0001;
}