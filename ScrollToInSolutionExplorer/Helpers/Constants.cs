namespace ScrollToInSolutionExplorer.Helpers
{
    /// <summary>
    /// General consts used to interact with Visual Studio.
    /// </summary>
    internal static class Constants
    {
        /// <summary>
        /// VS Command to locate and select the active document in Solution Explorer.
        /// </summary>
        public const string SyncWithActiveDocumentCommand = "SolutionExplorer.SyncWithActiveDocument";

        /// <summary>
        /// VS Command to collapse outlining in current active document.
        /// </summary>
        public const string CollapseToDefinitionCommand = "Edit.CollapsetoDefinitions";
    }
}
