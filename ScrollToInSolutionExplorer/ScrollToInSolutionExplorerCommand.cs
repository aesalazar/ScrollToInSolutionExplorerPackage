using System;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using ScrollToInSolutionExplorer.Helpers;
using Constants = ScrollToInSolutionExplorer.Helpers.Constants;

namespace ScrollToInSolutionExplorer
{
    /// <summary>
    /// Creates and registers the Commands used by the Menus and Toolbar.
    /// </summary>
    internal sealed class ScrollToInSolutionExplorerCommand
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrollToInSolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="disposalToken">Package disposal token.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        /// <param name="visualStudioInstance">Reference to the Visual Studio instance.</param>
        private ScrollToInSolutionExplorerCommand(
            CancellationToken disposalToken,
            OleMenuCommandService commandService,
            DTE2 visualStudioInstance)
        {
            _disposalToken = disposalToken;
            _visualStudioInstance = visualStudioInstance;

            _menuCommand = new OleMenuCommand(
                OnMenuCommandInvoke,
                null,
                OnMenuCommandBeforeQueryStatus,
                new CommandID(
                    PackageGuids.GuidScrollToInSolutionExplorerPackageCmdSet,
                    PackageIds.ScrollToInSolutionExplorerMenuCommandId));

            commandService.AddCommand(_menuCommand);

            _toolbarCommand = new OleMenuCommand(
                OnMenuCommandInvoke,
                null,
                OnMenuCommandBeforeQueryStatus,
                new CommandID(
                    PackageGuids.GuidScrollToInSolutionExplorerPackageCmdSet,
                    PackageIds.ScrollToInSolutionExplorerToolbarCommandId));

            commandService.AddCommand(_toolbarCommand);
        }

        #endregion

        #region Public Statics

        /// <summary>
        /// Initializes the singleton instance of the command that is registered with the command service.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="progress">Progress indicator to update.</param>
        /// <remarks>
        /// Call to AddCommand in ScrollToInSolutionExplorerCommand's constructor requires the UI thread.
        /// </remarks>
        public static async Task<ScrollToInSolutionExplorerCommand> InitializeAsync(AsyncPackage package, IProgress<ServiceProgressData> progress)
        {
            //Perform all async now so command can run synchronously
            progress.Report(new ServiceProgressData("Gathering components..."));
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync<IMenuCommandService, OleMenuCommandService>(package.DisposalToken);
            var applicationObject = await package.GetServiceAsync<_DTE, DTE2>(package.DisposalToken);

            progress.Report(new ServiceProgressData("Generating static instance..."));
            var command = new ScrollToInSolutionExplorerCommand(
                package.DisposalToken,
                commandService,
                applicationObject);

            progress.Report(new ServiceProgressData("Scroll To In Solution Explorer command initialized."));
            return command;
        }

        #endregion

        #region Fields

        /// <summary>
        /// Regisered Command
        /// </summary>
        private readonly OleMenuCommand _menuCommand;

        private readonly OleMenuCommand _toolbarCommand;

        /// <summary>
        /// Package disposal token.
        /// </summary>
        private readonly CancellationToken _disposalToken;

        /// <summary>
        /// Reference to Visual Studio.
        /// </summary>
        private readonly DTE2 _visualStudioInstance;

        #endregion

        #region Methods

        /// <summary>
        /// Attempts to locate the current active document in Solution Explorer and select it.
        /// </summary>
        /// <remarks>
        /// Seems more reliable when dispatched to a new task especially when mapped to a toolbar button.
        /// </remarks>
        [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "ExecuteTaskToUiThread dispatches to the UI Thread")]
        private void OnMenuCommandInvoke(object __, EventArgs ___)
        {
            _ = Task.Factory.StartNew(async () =>
            {
                try
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_disposalToken);
                    _visualStudioInstance
                        .DTE
                        .ExecuteCommand(Constants.SyncWithActiveDocumentCommand);
                }
                catch (Exception ex)
                {
                    if (ErrorHandler.IsCriticalException(ex))
                        throw;
                }
            }, _disposalToken, TaskCreationOptions.None, TaskScheduler.Default);
        }

        /// <summary>
        /// Determines if the command should be enabled or not based on current active document.
        /// </summary>
        /// <remarks>
        /// This cannot be dispatched in any way via a dispatcher or a new Task.  If the user has selected 
        /// a tab that is not selectable in Solution Explorer and then directly right-clicks on a file 
        /// that is, the icon in the context menu will not light up.  Seems the check in this method must 
        /// run synchronously so it sets the <see cref="MenuCommand.Enabled"/> before the context menu
        /// is rendered.
        /// </remarks>
        [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "ExecuteTaskToUiThread dispatches to the UI Thread")]
        private void OnMenuCommandBeforeQueryStatus(object _, EventArgs __)
        {
            try
            {
                var isEnabled = _visualStudioInstance
                    .DTE
                    .Commands
                    .Item(Constants.SyncWithActiveDocumentCommand)
                    .IsAvailable;

                _menuCommand.Enabled = isEnabled;
                _toolbarCommand.Enabled = isEnabled;
            }
            catch (ArgumentException)
            {
                _menuCommand.Enabled = false;
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;

                _menuCommand.Enabled = false;
            }
        }

        #endregion
    }
}