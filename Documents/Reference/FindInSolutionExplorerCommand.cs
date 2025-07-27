#nullable enable

using System;
using System.Linq;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System.Composition;
using System.ComponentModel.Design;

namespace ScrollToInSolutionExplorer
{
    internal class FindInSolutionExplorerCommand
    {
        private OleMenuCommand _command;

        public FindInSolutionExplorerCommand()
        {
            _command = new OleMenuCommand(
                OnMenuCommandInvoke, 
                OnMenuCommandChange, 
                OnMenuCommandBeforeQueryStatus,
                new CommandID(guidFindInSolutionExplorerCommandSet, cmdidFindInSolutionExplorer));

            var mcs = (IMenuCommandService)ServiceProvider.GetService(typeof(IMenuCommandService));
            mcs.AddCommand(_command);
        }

        public static readonly int cmdidFindInSolutionExplorer = 0x0100;

        public const string guidFindInSolutionExplorerCommandSetString = "B3AD9EAD-9439-445D-BECB-6176098247AC";
        public static readonly Guid guidFindInSolutionExplorerCommandSet = new Guid("{" + guidFindInSolutionExplorerCommandSetString + "}");

        public static EnvDTE80.Window2? FindWindow(EnvDTE80.Windows2 windows, EnvDTE.vsWindowType vsWindowType)
        {
            return windows
                .Cast<EnvDTE80.Window2>()
                .FirstOrDefault(w => w.Type == vsWindowType);
        }

        [Import]
        public SVsServiceProvider ServiceProvider { get; set; }

        public EnvDTE80.DTE2 ApplicationObject => (EnvDTE80.DTE2) ServiceProvider.GetService(typeof(EnvDTE._DTE));

        private void OnMenuCommandInvoke(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var track = ApplicationObject
                    .get_Properties("Environment", "ProjectsAndSolution")
                    .Item("TrackFileSelectionInExplorer");

                if (track.Value is bool && !((bool)track.Value))
                {
                    track.Value = true;
                    track.Value = false;
                }

                // Find the Solution Explorer object
                var windows = ApplicationObject.Windows as EnvDTE80.Windows2;
                var solutionExplorer = FindWindow(windows, EnvDTE.vsWindowType.vsWindowTypeSolutionExplorer);
                if (solutionExplorer != null)
                    solutionExplorer.Activate();
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
                EnvDTE.Document doc = ApplicationObject.ActiveDocument;

                _command.Supported = true;

                bool enabled = false;
                EnvDTE.ProjectItem projectItem = doc != null ? doc.ProjectItem : null;
                if (projectItem != null)
                {
                    if (projectItem.Document != null)
                    {
                        // normal project documents
                        enabled = true;
                    }
                    else if (projectItem.ContainingProject != null)
                    {
                        // this applies to files in the "Solution Files" folder
                        enabled = projectItem.ContainingProject.Object != null;
                    }
                }

                _command.Enabled = enabled;
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

    }
}
