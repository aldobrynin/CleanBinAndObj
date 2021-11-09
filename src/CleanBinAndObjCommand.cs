using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CleanBinAndObj
{
    /// <summary>
    ///     Command handler
    /// </summary>
    internal sealed class CleanBinAndObjCommand
    {
        /// <summary>
        ///     Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        ///     Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("58cab930-ec55-4b8b-876a-c6208cc246c4");

        /// <summary>
        ///     Pane to output command log messages
        /// </summary>
        private static IVsOutputWindowPane _vsOutputWindowPane;

        private readonly DTE2 _dte;
        private readonly Options _options;

        /// <summary>
        ///     VS Package that provides this command, not null.
        /// </summary>
        private readonly Package _package;

        private IVsStatusbar _bar;

        /// <summary>
        ///     Initializes a new instance of the <see cref="CleanBinAndObjCommand" /> class.
        ///     Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="dte2"></param>
        /// <param name="options"></param>
        private CleanBinAndObjCommand(Package package, DTE2 dte2, Options options)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _package = package ?? throw new ArgumentNullException(nameof(package));
            _dte = dte2;
            _options = options;

            if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                var menuCommandId = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(CleanBinAndObj, menuCommandId);
                commandService.AddCommand(menuItem);

                var outWindow = (IVsOutputWindow) Package.GetGlobalService(typeof(SVsOutputWindow));
                var generalPaneGuid =
                    VSConstants.GUID_BuildOutputWindowPane; // P.S. There's also the GUID_OutWindowDebugPane available.
                outWindow.GetPane(ref generalPaneGuid, out _vsOutputWindowPane);
                _vsOutputWindowPane.Activate(); // Brings this pane into view
            }
        }

        /// <summary>
        ///     Gets the instance of the command.
        /// </summary>
        public static CleanBinAndObjCommand Instance { get; private set; }

        /// <summary>
        ///     Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider => _package;

        /// <summary>
        ///     Gets the status bar.
        /// </summary>
        /// <value>The status bar.</value>
        private IVsStatusbar StatusBar => _bar ?? (_bar = ServiceProvider.GetService(typeof(SVsStatusbar)) as IVsStatusbar);

        /// <summary>
        ///     Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="dte"></param>
        /// <param name="options">options object</param>
        public static void Initialize(Package package, DTE2 dte, Options options)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Instance = new CleanBinAndObjCommand(package, dte, options);
        }

        private void CleanBinAndObj(object sender, EventArgs e)
        {
            var sw = new Stopwatch();
            var solutionProjects = GetProjects();
            _vsOutputWindowPane.Clear();
            WriteToOutput($"Starting... Projects to clean: {solutionProjects.Length}");
            uint cookie = 0;
            StatusBar.Progress(ref cookie, 1, string.Empty, 0, (uint) solutionProjects.Length);
            sw.Start();
            for (uint index = 0; index < solutionProjects.Length; index++)
            {
                var project = solutionProjects[index];
                var projectRootPath = GetProjectRootFolder(project);
                var message = $"Cleaning {project.UniqueName}";
                WriteToOutput(message);
                StatusBar.Progress(ref cookie, 1, string.Empty, index, (uint) solutionProjects.Length);
                StatusBar.SetText(message);
                CleanDirectory(projectRootPath);
            }

            sw.Stop();
            WriteToOutput($@"Finished. Elapsed: {sw.Elapsed:mm\:ss\.ffff}");
            // Clear the progress bar.
            StatusBar.Progress(ref cookie, 0, string.Empty, 0, 0);
            StatusBar.FreezeOutput(0);
            StatusBar.SetText("Cleaned bin and obj");
        }

        private void CleanDirectory(string directoryPath)
        {
            if (directoryPath == null)
                return;
            if (_options.TargetSubdirectories == null || _options.TargetSubdirectories.Length == 0) 
                return;
            try
            {
                foreach (var di in _options.TargetSubdirectories.Select(x => Path.Combine(directoryPath, x))
                    .Where(Directory.Exists)
                    .Select(x => new DirectoryInfo(x)))
                {
                    foreach (var file in di.EnumerateFiles())
                    {
                        file.Delete();
                    }

                    foreach (var dir in di.EnumerateDirectories())
                    {
                        dir.Delete(true);
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
        }

        private Project[] GetProjects()
        {
            var projects = GetActiveProjects(_dte) ?? _dte.Solution.Projects.Cast<Project>().ToArray();
            return projects
                .SelectMany(GetChildProjects)
                .Union(projects)
                .Where(ProjectFullNameNotEmpty)
                .OrderBy(x => x.UniqueName)
                .ToArray();
        }

        private static Project[] GetActiveProjects(DTE2 dte) {
            try {
                if (dte.ActiveSolutionProjects is Array activeSolutionProjects && activeSolutionProjects.Length > 0)
                {
                    return activeSolutionProjects.Cast<Project>().ToArray();
                }
            } catch (Exception ex) {
                Debug.Write(ex.Message);
            }

            return null;
        }

        private static bool ProjectFullNameNotEmpty(Project p)
        {
            try
            {
                return !string.IsNullOrEmpty(p.FullName);
            }
            catch
            {
                return false;
            }
        }

        private static IEnumerable<Project> GetChildProjects(Project parent)
        {
            try
            {
                if (parent.Kind != ProjectKinds.vsProjectKindSolutionFolder && parent.Collection == null) // Unloaded
                    return Enumerable.Empty<Project>();

                if (string.IsNullOrEmpty(parent.FullName) == false)
                    return Enumerable.Repeat(parent, 1);
            }
            catch (COMException)
            {
                return Enumerable.Empty<Project>();
            }

            return parent.ProjectItems
                .Cast<ProjectItem>()
                .Where(p => p.SubProject != null)
                .SelectMany(p => GetChildProjects(p.SubProject));
        }

        private static string GetProjectRootFolder(Project project)
        {
            if (string.IsNullOrEmpty(project.FullName))
                return null;

            string fullPath;

            try
            {
                fullPath = project.Properties.Item("FullPath").Value as string;
            }
            catch (ArgumentException)
            {
                try
                {
                    // MFC projects don't have FullPath, and there seems to be no way to query existence
                    fullPath = project.Properties.Item("ProjectDirectory").Value as string;
                }
                catch (ArgumentException)
                {
                    // Installer projects have a ProjectPath.
                    fullPath = project.Properties.Item("ProjectPath").Value as string;
                }
            }

            if (string.IsNullOrEmpty(fullPath))
                return File.Exists(project.FullName) ? Path.GetDirectoryName(project.FullName) : null;

            if (Directory.Exists(fullPath))
                return fullPath;

            return File.Exists(fullPath) ? Path.GetDirectoryName(fullPath) : null;
        }

        private void WriteToOutput(string message)
        {
            _vsOutputWindowPane.OutputStringThreadSafe($"{DateTime.Now:HH:mm:ss.ffff}: {message}{Environment.NewLine}");
        }
    }
}