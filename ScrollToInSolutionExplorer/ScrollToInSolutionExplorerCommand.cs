#nullable enable
using System;
using System.Collections;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Community.VisualStudio.Toolkit;
using EnvDTE;
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var applicationObject = (EnvDTE80.DTE2)await package.GetServiceAsync(typeof(_DTE));
            var commandService = (OleMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));
            Instance = new ScrollToInSolutionExplorerCommand(commandService, applicationObject, package);
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrollToInSolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="commandService">Command service to add command to, not null.</param>
        /// <param name="visualStudioInstance">Reference to the Visual Studio instance.</param>
        private ScrollToInSolutionExplorerCommand(
            OleMenuCommandService commandService,
            EnvDTE80.DTE2 visualStudioInstance,
            AsyncPackage package)
        {
            _visualStudioInstance = visualStudioInstance;
            _package = package;
            _command = new OleMenuCommand(
                OnMenuCommandInvoke,
                OnMenuCommandChange,
                OnMenuCommandBeforeQueryStatus,
                new CommandID(CommandSet, CommandId)
            );

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
        private OleMenuCommand _command;

        /// <summary>
        /// Reference to Visual Studio.
        /// </summary>
        private EnvDTE80.DTE2 _visualStudioInstance;

        private readonly AsyncPackage _package;

        #endregion

        #region Methods

        private static EnvDTE80.Window2? FindWindow(EnvDTE80.Windows2 windows, EnvDTE.vsWindowType vsWindowType)
        {
            return windows
                .Cast<EnvDTE80.Window2>()
                .FirstOrDefault(w => w.Type == vsWindowType);
        }

        #endregion

        #region Command Handlers

        private void OnMenuCommandInvoke(object sender, EventArgs e)
        {
            _ = Task.Run(() => OnMenuCommandInvokeAsync(sender, e));
        }


        private async Task SelectCurrentDocumentAsync()
        {
            var docView = await VS.Documents.GetActiveDocumentViewAsync();

            if (docView?.FilePath != null)
            {
                var item = await PhysicalFile.FromFileAsync(docView.FilePath);

                // Find the Solution Explorer object
                var windows = (EnvDTE80.Windows2)_visualStudioInstance.Windows;
                var solutionExplorer = FindWindow(windows, vsWindowType.vsWindowTypeSolutionExplorer);
                if (solutionExplorer != null)
                {
                    Debug.WriteLine($"INFO: SE Document = {solutionExplorer.Document?.Name}");
                    solutionExplorer.Set()
                }


            }
        }

        private async Task<string?> SelectedItemFilePathAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var uiHeirarcy = _visualStudioInstance.ToolWindows.SolutionExplorer;

            if (uiHeirarcy.SelectedItems is IEnumerable selectedItems)
            {
                foreach (UIHierarchyItem selItem in selectedItems)
                {
                    if (selItem.Object is ProjectItem projectItem)
                    {
                        Debug.WriteLine($"INFO: prjItem = {projectItem.Name}");
                        var path = projectItem.Properties.Item("FullPath").Value.ToString(); ;
                        var item = await PhysicalFile.FromFileAsync(path);
                        Debug.WriteLine($"INFO: item = {item?.FullPath}");
                    }
                }

                foreach (UIHierarchyItem selItem in selectedItems)
                {
                    if (selItem.Object is ProjectItem prjItem)
                        return prjItem.Properties.Item("FullPath").Value.ToString();
                }
            }

            return null;
        }

        private async Task<string?> GetSolutionFilterFilePathAsync()
        {
            Community.VisualStudio.Toolkit.Solution? solution;
            solution = await VS.Solutions.GetCurrentSolutionAsync();

            if (solution is not null)
            {
                solution.GetItemInfo(out IVsHierarchy hierarchy, out _, out _);
                if (hierarchy.GetType().GetProperty("SolutionFilterFilePath") is PropertyInfo filePath)
                {
                    return filePath.GetValue(hierarchy) as string;
                }
            }

            return null;
        }

        private async Task OnMenuCommandInvokeAsync(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            try
            {
                Debug.WriteLine($"INFO: SelectedItemFilePathAsync = {await SelectedItemFilePathAsync()}");

                //Use the VS track file selection to actually select
                var trackFileProperty = _visualStudioInstance
                    .Properties["Environment", "ProjectsAndSolution"]
                    .Item("TrackFileSelectionInExplorer");

                //Activate if currently off
                var isTrackingSelection = trackFileProperty.Value is true;
                if (!isTrackingSelection)
                {
                    trackFileProperty.Value = true;
                    await Task.Delay(500).ConfigureAwait(true); // YIKES!!!!
                    trackFileProperty.DTE.MainWindow.Activate();
                    Debug.WriteLine($"INFO: Track Value 1 = {trackFileProperty.Value}");
                }

                await Task.Factory.StartNew(async () =>
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    try
                    {
                        // Find the Solution Explorer object
                        var windows = (EnvDTE80.Windows2)_visualStudioInstance.Windows;
                        var solutionExplorer = FindWindow(windows, vsWindowType.vsWindowTypeSolutionExplorer);
                        if (solutionExplorer != null)
                        {
                            Debug.WriteLine($"INFO: SE Document = {solutionExplorer.Document?.Name}");
                            solutionExplorer.Activate();
                        }
                    }
                    finally
                    {
                        //Restore the prior state if neccessary
                        if (!isTrackingSelection)
                            await ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                            {
                                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                                trackFileProperty.Value = false;
                                Debug.WriteLine($"INFO: Track Value 2 = {trackFileProperty.Value}");
                            });
                    }
                }, CancellationToken.None, TaskCreationOptions.RunContinuationsAsynchronously, TaskScheduler.Default);
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;
            }
        }

        private void OnMenuCommandChange(object sender, EventArgs e)
        {
        }

        private void OnMenuCommandBeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var doc = _visualStudioInstance.ActiveDocument;
                _command.Supported = true;

                var isCommandEnabled = false;
                var projectItem = doc?.ProjectItem;
                if (projectItem != null)
                {
                    Debug.WriteLine($"INFO: Item = {projectItem.Name} - {projectItem.Kind}");

                    if (projectItem.Document != null)
                    {
                        // normal project documents
                        isCommandEnabled = true;
                    }
                    else if (projectItem.ContainingProject != null)
                    {
                        // this applies to files in the "Solution Files" folder
                        isCommandEnabled = projectItem.ContainingProject.Object != null;
                    }
                }

                _command.Enabled = isCommandEnabled;
            }
            catch (ArgumentException)
            {
                // stupid thing throws if the active window is a C# project properties pane
                _command.Supported = false;
                _command.Enabled = false;
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;

                _command.Supported = false;
                _command.Enabled = false;
            }
        }

        #endregion
    }
}
