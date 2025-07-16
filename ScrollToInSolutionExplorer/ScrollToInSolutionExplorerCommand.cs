#nullable enable
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

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
        public static async Task InitializeAsync(AsyncPackage package, IProgress<ServiceProgressData> progress)
        {
            //Perform all async now so command can run synchronously
            progress.Report(new ServiceProgressData("Gathering components..."));
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = (OleMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));
            var applicationObject = (DTE2)await package.GetServiceAsync(typeof(_DTE));
            var vsSolutionHierarchy = (IVsHierarchy)await package.GetServiceAsync(typeof(SVsSolution));

            progress.Report(new ServiceProgressData("Generating static instance..."));
            var command = new ScrollToInSolutionExplorerCommand(
                package.DisposalToken,
                commandService,
                applicationObject,
                vsSolutionHierarchy);

            progress.Report(new ServiceProgressData("Scroll To In Solution Explorer command initialized."));
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
        /// Seems more reliable when dispatched to a new task especially when mapped to a tool bar button.
        /// </remarks>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "ExecuteTaskToUiThread dispatches to the UI Thread")]
        private void OnMenuCommandInvoke(object _, EventArgs __)
        {
            _ = Task.Factory.StartNew(async () =>
            {
                try
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_disposalToken);
                    _visualStudioInstance.DTE.ExecuteCommand(Constants.SyncWithActiveDocumentCommand);
                }
                catch (Exception ex)
                {
                    if (ErrorHandler.IsCriticalException(ex))
                        throw;
                }
            }, _disposalToken, TaskCreationOptions.None, TaskScheduler.Default);
        }

        /// <summary>
        /// Determines if the command should enabled or not based on current active document.
        /// </summary>
        /// <remarks>
        /// This cannot be dispatched in any way via a dispatcher or a new Task.  If the user has selected 
        /// a tab that is not selectable in Solution Explorer and then directly right-clicks on a file 
        /// that is, the icon in the context menu will not light up.  Seems the check in this method must 
        /// run synchronously so it sets the <see cref="MenuCommand.Enabled"/> before the context menu
        /// is rendered.
        /// </remarks>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "ExecuteTaskToUiThread dispatches to the UI Thread")]
        private void OnMenuCommandBeforeQueryStatus(object _, EventArgs __)
        {
            try
            {
                var visibleName = _visualStudioInstance.ActiveWindow.Project?.Name ?? _visualStudioInstance.ActiveWindow.Caption;
                _command.Enabled = visibleName is string name && SolutionExplorerHelpers.IsDisplayNameInHierarcy(_vsSolutionHierarchy, name);
                Debug.WriteLine($"INFO: display name '{_visualStudioInstance.ActiveWindow.Project?.Name} - {_visualStudioInstance.ActiveWindow.Caption}' present: {_command.Enabled}");
            }
            catch (ArgumentException)
            {
                _command.Enabled = false;
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;

                _command.Enabled = false;
            }
        }

        #endregion
    }
}
