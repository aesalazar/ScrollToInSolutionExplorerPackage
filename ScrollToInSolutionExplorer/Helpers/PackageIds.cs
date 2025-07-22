namespace ScrollToInSolutionExplorer.Helpers;

/// <summary>
/// Helper class that encapsulates all CommandIDs uses across VS Package.
/// </summary>
/// <remarks>
/// Names here are based on the XML in the .vsct file.
/// </remarks>
internal static class PackageIds
{
    public const int ScrollToInSolutionExplorerMenuGroup = 0x1020;
    public const int ScrollToInSolutionExplorerMenuCommandId = 0x0100;

    public const int ScrollToInSolutionExplorerToolbar = 0x1031;
    public const int ScrollToInSolutionExplorerToolbarCommandId = 0x1032;

    public const int BmpScrollTo = 0x0001;
}