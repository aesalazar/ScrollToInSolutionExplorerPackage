using System;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using ScrollToInSolutionExplorer.Helpers;

namespace ScrollToInSolutionExplorer;

/// <summary>
/// Creates and registers the Commands used by the Menus and Toolbar.
/// </summary>
internal sealed class ScrollToInSolutionExplorerCommandManager
{
    #region Constructor

    /// <summary>
    /// Initializes a new instance of the class.
    /// </summary>
    /// <param name="package">Parent Package.</param>
    /// <param name="commandService">Command service to add command to, not null.</param>
    /// <param name="visualStudioInstance">Reference to the Visual Studio instance.</param>
    private ScrollToInSolutionExplorerCommandManager(
        AsyncPackage package,
        OleMenuCommandService commandService,
        DTE2 visualStudioInstance)
    {
        _package = package;
        _visualStudioInstance = visualStudioInstance;

        _menuCommand = new OleMenuCommand(
            OnCommandInvoke,
            null,
            OnCommandBeforeQueryStatus,
            new CommandID(
                PackageGuids.GuidScrollToInSolutionExplorerPackageCmdSet,
                PackageIds.ScrollToInSolutionExplorerMenuCommandId));

        commandService.AddCommand(_menuCommand);

        _tabMenuCommand = new OleMenuCommand(
            OnCommandInvoke,
            null,
            OnCommandBeforeQueryStatus,
            new CommandID(
                PackageGuids.GuidScrollToInSolutionExplorerPackageCmdSet,
                PackageIds.ScrollToInSolutionExplorerDocumentTabCommandId));

        commandService.AddCommand(_tabMenuCommand);

        _toolbarCommand = new OleMenuCommand(
            OnCommandInvoke,
            null,
            OnCommandBeforeQueryStatus,
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
    public static async Task<ScrollToInSolutionExplorerCommandManager> InitializeAsync(
        AsyncPackage package,
        IProgress<ServiceProgressData> progress)
    {
        //Perform all async now so command can run synchronously
        progress.Report(new ServiceProgressData("Gathering components..."));
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
        var commandService = await package.GetServiceAsync<IMenuCommandService, OleMenuCommandService>(package.DisposalToken);
        var applicationObject = await package.GetServiceAsync<_DTE, DTE2>(package.DisposalToken);

        progress.Report(new ServiceProgressData("Generating static instance..."));
        var manager = new ScrollToInSolutionExplorerCommandManager(
            package,
            commandService,
            applicationObject);

        progress.Report(new ServiceProgressData("Scroll To In Solution Explorer command initialized."));
        return manager;
    }

    #endregion

    #region Fields

    /// <summary>
    /// Registered Command for the Visual Studio Menu.
    /// </summary>
    private readonly OleMenuCommand _menuCommand;

    /// <summary>
    /// Registered command for the Document Tab Menu.
    /// </summary>
    private readonly OleMenuCommand _tabMenuCommand;

    /// <summary>
    /// Registered command for the custom Toolbar.
    /// </summary>
    private readonly OleMenuCommand _toolbarCommand;

    /// <summary>
    /// Parent Package.
    /// </summary>
    private readonly AsyncPackage _package;

    /// <summary>
    /// Reference to Visual Studio.
    /// </summary>
    private readonly DTE2 _visualStudioInstance;

    #endregion

    #region Methods

    /// <summary>
    /// Updates state of UI controls based on the latest options.
    /// </summary>
    /// <param name="optionsPage">Use-specified option settings.</param>
    public void ApplyUserSettings(ScrollToInSolutionExplorerOptionsPage optionsPage)
    {
        _menuCommand.Visible = optionsPage.IsVisibleInToolMenu;
        _menuCommand.Supported = optionsPage.IsVisibleInToolMenu;

        _tabMenuCommand.Visible = optionsPage.IsVisibleInDocumentTabMenu;
        _tabMenuCommand.Supported = optionsPage.IsVisibleInDocumentTabMenu;
    }

    /// <summary>
    /// Attempts to locate the current active document in Solution Explorer and select it.
    /// </summary>
    /// <remarks>
    /// Seems more reliable when dispatched to a new task especially when mapped to a toolbar button.
    /// </remarks>
    [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "ExecuteTaskToUiThread dispatches to the UI Thread")]
    private void OnCommandInvoke(object __, EventArgs ___)
    {
        _ = Task.Factory.StartNew(async () =>
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);
                _visualStudioInstance
                    .DTE
                    .ExecuteCommand(GeneralConstants.SyncWithActiveDocumentCommand);
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;
            }
        }, _package.DisposalToken, TaskCreationOptions.None, TaskScheduler.Default);
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
    private void OnCommandBeforeQueryStatus(object _, EventArgs __)
    {
        try
        {
            var isEnabled = _visualStudioInstance
                .DTE
                .Commands
                .Item(GeneralConstants.SyncWithActiveDocumentCommand)
                .IsAvailable;

            _menuCommand.Enabled = isEnabled;
            _toolbarCommand.Enabled = isEnabled;
            _tabMenuCommand.Enabled = isEnabled;
        }
        catch (ArgumentException)
        {
            _menuCommand.Enabled = false;
            _tabMenuCommand.Enabled = false;
            _toolbarCommand.Enabled = false;
        }
        catch (Exception ex)
        {
            if (ErrorHandler.IsCriticalException(ex))
                throw;

            _menuCommand.Enabled = false;
            _tabMenuCommand.Enabled = false;
            _toolbarCommand.Enabled = false;
        }
    }

    #endregion
}