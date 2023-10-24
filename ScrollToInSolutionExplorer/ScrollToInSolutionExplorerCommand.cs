#nullable enable
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace ScrollToInSolutionExplorer
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ScrollToInSolutionExplorerCommand
    {
        #region Public Statics

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ScrollToInSolutionExplorerCommand? Instance { get; private set; }

        /// <summary>
        /// Initializes the singleton instance of the command that is registered with the command service.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <remarks>
        /// Call to AddCommand in ScrollToInSolutionExplorerCommand's constructor requires the UI thread.
        /// </remarks>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            //Perform all async now so command can run synchronously
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = (OleMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));
            var applicationObject = (DTE2)await package.GetServiceAsync(typeof(_DTE));
            var vsSolutionHierarchy = (IVsHierarchy)await package.GetServiceAsync(typeof(SVsSolution));

            Instance = new ScrollToInSolutionExplorerCommand(
                package.DisposalToken,
                commandService,
                applicationObject,
                vsSolutionHierarchy);
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrollToInSolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="commandService">Command service to add command to, not null.</param>
        /// <param name="visualStudioInstance">Reference to the Visual Studio instance.</param>
        /// <param name="vsSolutionHierarchy">Reference to the Visual Stuio root hierarchy.</param>
        private ScrollToInSolutionExplorerCommand(
            CancellationToken disposalToken,
            OleMenuCommandService commandService,
            DTE2 visualStudioInstance,
            IVsHierarchy vsSolutionHierarchy)
        {
            _disposalToken = disposalToken;
            _visualStudioInstance = visualStudioInstance;
            _vsSolutionHierarchy = vsSolutionHierarchy;

            _command = new OleMenuCommand(
                OnMenuCommandInvoke,
                null,
                OnMenuCommandBeforeQueryStatus,
                new CommandID(CommandSet, CommandId));

            commandService.AddCommand(_command);
        }

        #endregion

        #region Fields

        /// <summary>
        /// Command ID.
        /// </summary>
        private const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        private static readonly Guid CommandSet = new Guid("d41f8793-ad72-4ed1-af0f-cec3de2b9a61");

        /// <summary>
        /// Regisered Command
        /// </summary>
        private readonly OleMenuCommand _command;

        /// <summary>
        /// Package disposal token.
        /// </summary>
        private readonly CancellationToken _disposalToken;

        /// <summary>
        /// Reference to Visual Studio.
        /// </summary>
        private readonly DTE2 _visualStudioInstance;

        /// <summary>
        /// Reference to the Visual Stuio root hierarchy.
        /// </summary>
        private readonly IVsHierarchy _vsSolutionHierarchy;

        #endregion

        #region Methods

        /// <summary>
        /// Attempts to locate the current active document in Solution Explorer and select it.
        /// </summary>
        /// <remarks>
        /// Runnin as async causes selection to fire much more slowly.
        /// </remarks>
        private void OnMenuCommandInvoke(object _, EventArgs __)
        {
            if (!_command.Enabled)
                return;

            try
            {
                //Total hack - seems to have problem navigating to files at the bottom of the tree so go to the top first
                ThreadHelper.ThrowIfNotOnUIThread();
                var nodeNames = SolutionExplorerHelpers.SelectRootInSolutionExplorer(_visualStudioInstance, _vsSolutionHierarchy);
                Debug.WriteLine($"INFO: Node Path {string.Join("->", nodeNames)}");

                _ = Task.Factory.StartNew(async () =>
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    var activeDocument = _visualStudioInstance.ActiveDocument;
                    if (activeDocument is null)
                        return;

                    var fileName = activeDocument.FullName;
                    Debug.WriteLine($"INFO: CURRENT DOCUMENT:  {fileName}");
                    if (fileName is null)
                        return;

                    var nodeNames = SolutionExplorerHelpers.SelectInSolutionExplorer(
                        _visualStudioInstance,
                        _vsSolutionHierarchy,
                        fileName);

                    Debug.WriteLine($"INFO: Node Path {string.Join("->", nodeNames)}");
                }, _disposalToken, TaskCreationOptions.None, TaskScheduler.Default);
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;
            }
        }

        /// <summary>
        /// Determines if the command should enabled or not based on current active document.
        /// </summary>
        /// <remarks>
        /// This cannot be async as the commands will not enable/disable in time in the UI.
        /// </remarks>
        private void OnMenuCommandBeforeQueryStatus(object _, EventArgs __)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var isCommandEnabled = false;

                //First check for regular project items
                var projectItem = _visualStudioInstance.ActiveDocument?.ProjectItem;
                if (projectItem is not null)
                {
                    Debug.WriteLine($"INFO: Item = {projectItem.Name} - {projectItem.Kind}");

                    if (projectItem.Document != null)
                        // normal project documents
                        isCommandEnabled = true;
                    else if (projectItem.ContainingProject != null)
                        // this applies to files in the "Solution Files" folder
                        isCommandEnabled = projectItem.ContainingProject.Object != null;
                }
                else
                {
                    var caption = _visualStudioInstance.ActiveWindow.Caption;
                    if (caption is not null)
                        isCommandEnabled = true;
                }

                _command.Enabled = isCommandEnabled;
            }
            catch (ArgumentException ex)
            {
                Debug.WriteLine($"WARNING: Scroll To Item not supported: {ex.Message}");
                _command.Enabled = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ERROR: Unexpected exception: {ex}");
                if (ErrorHandler.IsCriticalException(ex))
                    throw;

                _command.Enabled = false;
            }
        }

        #endregion
    }
}
