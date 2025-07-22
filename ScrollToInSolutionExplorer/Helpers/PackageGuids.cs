using System;

namespace ScrollToInSolutionExplorer.Helpers;

/// <summary>
/// Helper class that exposes all GUIDs used across VS Package.
/// </summary>
/// <remarks>
/// Names here are based on the XML in the .vsct file.
/// </remarks>
internal static class PackageGuids
{
    public const string GuidScrollToInSolutionExplorerPackageString = "ba7408f0-ed7d-4efa-85ef-2d648b9a3274";
    public const string GuidScrollToInSolutionExplorerPackageCmdSetString = "d41f8793-ad72-4ed1-af0f-cec3de2b9a61";
    public const string GuidImagesString = "3ef7db1b-bdbc-4232-9d89-9f8a8503e7e3";
    
    public static Guid GuidScrollToInSolutionExplorerPackage = new(GuidScrollToInSolutionExplorerPackageString);
    public static Guid GuidScrollToInSolutionExplorerPackageCmdSet = new(GuidScrollToInSolutionExplorerPackageCmdSetString);
    public static Guid GuidImages = new(GuidImagesString);
}