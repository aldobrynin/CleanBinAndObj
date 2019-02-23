using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace CleanBinAndObj
{

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(Vsix.Id)]
    [ProvideOptionPage(typeof(Options), "Environment", Vsix.Name, 101, 102, true, new string[] { }, ProvidesLocalizedCategoryName = false)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class CleanBinAndObjCommandPackage : Package
    {

        protected override void Initialize()
        {
            var dte = GetService(typeof(DTE)) as DTE2;
            var options = (Options)GetDialogPage(typeof(Options));

            CleanBinAndObjCommand.Initialize(this, dte, options);
            base.Initialize();
        }
    }
}
