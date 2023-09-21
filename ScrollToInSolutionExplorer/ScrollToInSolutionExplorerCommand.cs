using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Community.VisualStudio.Toolkit;
using Microsoft.VisualStudio.Shell;

namespace ScrollToInSolutionExplorer
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ScrollToInSolutionExplorerCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("d41f8793-ad72-4ed1-af0f-cec3de2b9a61");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrollToInSolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ScrollToInSolutionExplorerCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ScrollToInSolutionExplorerCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ScrollToInSolutionExplorerCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ScrollToInSolutionExplorerCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            _ = Task.Run(() => ExecuteAsync(sender, e));
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async Task ExecuteAsync(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            //Look for a match by path
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            var path = documentView?.FilePath;
            Debug.WriteLine($"CURRENT DOCUMENT:  {path}");
            if (string.IsNullOrEmpty(path))
                return;

            //Look for the root node
            var solutionExplorer = await VS.Windows.GetSolutionExplorerWindowAsync();
            var selections = await solutionExplorer.GetSelectionAsync();
            if (!selections.Any())
                return;

            var rootNode = selections.First();
            while (rootNode.Parent != null)
            {
                rootNode = rootNode.Parent;
            }

            //Find a match
            var match = default(SolutionItem);
            SolutionExplorerHelpers.TraverseChildren(
                rootNode
                , solutionItem =>
                {
                    Debug.WriteLine($"AVAILABLE DOCUMENT: {solutionItem.Type}: {solutionItem.Name}");
                    if (solutionItem.FullPath != path)
                        return false;

                    match = solutionItem;
                    return true;
                });

            if (match == default)
                return;

            solutionExplorer.SetSelection(match);
            Debug.WriteLine($"Selected file in Solution Explorer: {path}");
        }
    }
}
