using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace CleanBinAndObj
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(Vsix.Id)]
    [ProvideOptionPage(typeof(Options), "Environment", Vsix.Name, 101, 102, true, new string[] { }, ProvidesLocalizedCategoryName = false)]
    public sealed class CleanBinAndObjPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            var options = (Options)GetDialogPage(typeof(Options));

            CleanBinAndObjCommand.Initialize(this, dte, options);
        }
    }
}
