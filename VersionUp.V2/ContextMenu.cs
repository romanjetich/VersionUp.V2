using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;

namespace VersionUp.V2
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ContextMenu
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandMajor = 0x0099;
        public const int CommandMinor = 0x0100;
        public const int CommandBuild = 0x0200;
        public const int CommandRevision = 0x0300;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a8031b27-b8a1-4683-9f1e-b1b86f6f237a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ContextMenu"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ContextMenu(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;
        }

        public async Task InitializeMenuAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;

            if (commandService != null)
            {
                var menuCommandMajor = new CommandID(CommandSet, CommandMajor);
                var menuMajor = new MenuCommand(this.MenuMajorCallback, menuCommandMajor);
                commandService.AddCommand(menuMajor);

                var menuCommandMinor = new CommandID(CommandSet, CommandMinor);
                var menuMinor = new MenuCommand(this.MenuMinorCallback, menuCommandMinor);
                commandService.AddCommand(menuMinor);

                var menuCommandBuild = new CommandID(CommandSet, CommandBuild);
                var menuBuild = new MenuCommand(this.MenuBuildCallback, menuCommandBuild);
                commandService.AddCommand(menuBuild);

                var menuCommandRevision = new CommandID(CommandSet, CommandRevision);
                var menuRevision = new MenuCommand(this.MenuRevisionCallback, menuCommandRevision);
                commandService.AddCommand(menuRevision);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ContextMenu Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void InitializeSingleton(Package package)
        {
            Instance = new ContextMenu(package);
        }

        private string GetActiveFilePath(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE2 applicationObject = serviceProvider.GetService(typeof(DTE)) as DTE2;
            if (applicationObject == null) throw new Exception("Unable access the service object");
            return applicationObject.ActiveDocument.FullName;
        }

        public enum Part
        {
            Major,
            Minor,
            Build,
            Revision,
        }

        /// <summary>
        /// Updates the assembly version
        /// </summary>
        private void UpVersion(Part part)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Project selProject = GetSelectedProject();
            string title = "Version Up";
            string message = "Please select the project!";
            Version currentVersion;
            string newAssemblyVersion = "";

            if (selProject != null)
            {
                int count = selProject.Properties.Count;
                try
                {
                    var sourceVersion = selProject.Properties.Item("Version").Value.ToString();
                    var parts = sourceVersion.Split('.').Length;
                    currentVersion = new Version(sourceVersion);
                    //var PackageTags = selProject.Properties.Item("PackageTags").Value;
                    //var Product = selProject.Properties.Item("Product").Value;
                    //var LocalPath = selProject.Properties.Item("LocalPath").Value;
                    //var SupportedTargetFrameworks = selProject.Properties.Item("SupportedTargetFrameworks").Value;
                    //var FullPath = selProject.Properties.Item("FullPath").Value;
                    //var Version = selProject.Properties.Item("Version").Value;
                    //var TargetFrameworkMoniker = selProject.Properties.Item("TargetFrameworkMoniker").Value;
                    //var FileVersion = selProject.Properties.Item("FileVersion").Value;
                    //var TargetFramework = selProject.Properties.Item("TargetFramework").Value;
                    //var TargetFrameworks = selProject.Properties.Item("TargetFrameworks").Value;

                    var major = currentVersion.Major;
                    var minor = currentVersion.Minor;
                    var build = currentVersion.Build;
                    var revision = currentVersion.Revision;

                    switch (part)
                    {
                        case Part.Major:
                            if (parts < 1) parts = 1;

                            major++;
                            minor = 0;
                            build = 0;
                            revision = 0;

                            break;
                        case Part.Minor:
                            if (parts < 2) parts = 2;

                            if (major < 0) major = 0;
                            minor++;
                            build = 0;
                            revision = 0;

                            break;
                        case Part.Build:
                            if (parts < 3) parts = 3;

                            if (major < 0) major = 0;
                            if (minor < 0) minor = 0;
                            build++;
                            revision = 0;

                            break;
                        case Part.Revision:
                            if (parts < 4) parts = 4;

                            if (major < 0) major = 0;
                            if (minor < 0) minor = 0;
                            if (build < 0) build = 0;
                            revision++;

                            break;
                        default:
                            throw new ArgumentOutOfRangeException(nameof(part), part, null);
                    }

                    switch (parts)
                    {
                        case 1: newAssemblyVersion = $"{major}"; break;
                        case 2: newAssemblyVersion = $"{major}.{minor}"; break;
                        case 3: newAssemblyVersion = $"{major}.{minor}.{build}"; break;
                        case 4:
                        default:
                            newAssemblyVersion = $"{major}.{minor}.{build}.{revision}"; break;
                    }

                    //if (selProject.Properties.Item("AssemblyVersion").Value != null)
                    //    selProject.Properties.Item("AssemblyVersion").Value = newAssemblyVersion;
                    //try
                    //{
                    //    if (selProject.Properties.Item("AssemblyFileVersion").Value != null)
                    //        selProject.Properties.Item("AssemblyFileVersion").Value = newAssemblyVersion;
                    //}
                    //catch (Exception e){}

                    // .NetCore Project doesn't has a property AssemblyFileVersion
                    try
                    {
                        //if (selProject.Properties.Item("FileVersion").Value != null)
                        //    selProject.Properties.Item("FileVersion").Value = newAssemblyVersion;
                        if (selProject.Properties.Item("Version").Value != null)
                            selProject.Properties.Item("Version").Value = newAssemblyVersion; // Package Version
                    }
                    catch (Exception e) { }

                    message = selProject.Name + " " + sourceVersion + " -> " + newAssemblyVersion;
                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox(
                        this.ServiceProvider,
                        message,
                        title,
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
                catch (Exception e)
                {
                    VsShellUtilities.ShowMessageBox(
                    this.ServiceProvider,
                    message,
                    e.Message,
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            }
        }

        /// <summary>
        /// Get the selected project in the solution.
        /// </summary>
        /// <returns>Selected project or null.</returns>
        private Project GetSelectedProject()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            IntPtr hierarchyPointer, selectionContainerPointer;
            Object selectedObject = null;
            IVsMultiItemSelect multiItemSelect;
            uint projectItemId;

            IVsMonitorSelection monitorSelection =
                    (IVsMonitorSelection)Package.GetGlobalService(
                    typeof(SVsShellMonitorSelection));

            monitorSelection.GetCurrentSelection(out hierarchyPointer,
                                                 out projectItemId,
                                                 out multiItemSelect,
                                                 out selectionContainerPointer);

            IVsHierarchy selectedHierarchy = null;
            try
            {
                selectedHierarchy = Marshal.GetTypedObjectForIUnknown(
                                                     hierarchyPointer,
                                                     typeof(IVsHierarchy)) as IVsHierarchy;
            }
            catch (Exception)
            {
                return null;
            }

            if (selectedHierarchy != null)
            {
                ErrorHandler.ThrowOnFailure(selectedHierarchy.GetProperty(
                                                  projectItemId,
                                                  (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                  out selectedObject));
            }

            Project selectedProject = selectedObject as Project;
            return selectedProject;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuMajorCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(Part.Major);
        }

        private void MenuMinorCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(Part.Minor);
        }

        private void MenuBuildCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(Part.Build);
        }

        private void MenuRevisionCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(Part.Revision);
        }
    }
}
