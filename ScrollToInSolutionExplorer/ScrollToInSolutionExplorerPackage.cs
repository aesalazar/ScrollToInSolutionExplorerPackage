using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ScrollToInSolutionExplorer.Helpers;
using Task = System.Threading.Tasks.Task;

namespace ScrollToInSolutionExplorer;

/// <summary>
/// This is the class that implements the package exposed by this assembly.
/// </summary>
[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
[Guid(PackageGuids.GuidScrollToInSolutionExplorerPackageString)]
[ProvideMenuResource("Menus.ctmenu", 1)]
[ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
[ProvideOptionPage(typeof(ScrollToInSolutionExplorerOptionsPage), GeneralConstants.ProjectName, GeneralConstants.OptionsPageHeader, 0, 0, true)]
public sealed class ScrollToInSolutionExplorerPackage : AsyncPackage
{
    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited, so this is the place
    /// where you can put all the initialization code that rely on services provided by VisualStudio.
    /// </summary>
    /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
    /// <param name="progress">A provider for progress updates.</param>
    /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
    /// <remarks>
    /// When initialized asynchronously, the current thread may be a background thread at this point.  Do any 
    /// initialization that requires the UI thread after switching to the UI thread.
    /// </remarks>
    protected override async Task InitializeAsync(
        CancellationToken cancellationToken,
        IProgress<ServiceProgressData> progress)
    {
        const string logName = GeneralConstants.ProjectName;

        progress.Report(new ServiceProgressData($"{logName} switching to UI Thread."));
        await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

        progress.Report(new ServiceProgressData($"{logName} initializing Command Manager."));
        var manager = await ScrollToInSolutionExplorerCommandManager.InitializeAsync(this, progress);

        progress.Report(new ServiceProgressData($"{logName} Initializing Options Dialog."));
        var options = (ScrollToInSolutionExplorerOptionsPage)GetDialogPage(typeof(ScrollToInSolutionExplorerOptionsPage));
        options.Initialize(manager);

        progress.Report(new ServiceProgressData("Initializing Scroll To In Solution Explorer complete."));
    }
}