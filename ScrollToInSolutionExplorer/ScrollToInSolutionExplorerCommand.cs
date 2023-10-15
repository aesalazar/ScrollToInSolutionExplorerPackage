#nullable enable
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var applicationObject = (DTE2)await package.GetServiceAsync(typeof(_DTE));
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
            DTE2 visualStudioInstance,
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
        private readonly OleMenuCommand _command;

        /// <summary>
        /// Reference to Visual Studio.
        /// </summary>
        private readonly DTE2 _visualStudioInstance;

        /// <summary>
        /// Reference to the contain extension package.
        /// </summary>
        private readonly AsyncPackage _package;

        #endregion

        #region Methods

        private void OnMenuCommandInvoke(object sender, EventArgs e)
        {
            _ = Task.Run(() => OnMenuCommandInvokeAsync(sender, e));
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
                _command.Supported = true;
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

        private void OnMenuCommandChange(object sender, EventArgs e)
        {
        }

        private async Task OnMenuCommandInvokeAsync(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);
            var activeDocument = _visualStudioInstance.ActiveDocument;
            if(activeDocument is null)
                return;

            var fileName = activeDocument.Name;
            Debug.WriteLine($"INFO: CURRENT DOCUMENT:  {fileName}");
            if (fileName is null)
                return;

            //Get the VS soluton as its hierarchy
            var vsSolutionHierarchy = await _package.GetServiceAsync(typeof(SVsSolution)) as IVsHierarchy;
            if (vsSolutionHierarchy is null)
                return;

            var nodeNames = new List<string>();
            await SolutionExplorerHelpers.FindItemInHierarchyAsync(
                _package,
                vsSolutionHierarchy,
                Path.GetFileName(fileName)
            );

            Debug.WriteLine($"INFO: Node Path {string.Join("->", nodeNames)}");
        }

        #endregion
    }
}
