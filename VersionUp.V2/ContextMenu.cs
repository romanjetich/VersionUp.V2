using System;
//using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;
using System.ComponentModel.Design;

namespace VersionUp.V2
{
    /// <summary>
    /// Command handler
    /// </summary>
    [GuidAttribute("9ED54F84-A89D-4fcd-A854-44251E925F09")]
    internal sealed class ContextMenu
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandMajor = 0x0100;
        public const int CommandMinor = 0x0200;
        public const int CommandBuild = 0x0300;
        public const int CommandRevision = 0x0400;

        public enum VersionTypes
        {
            Major,
            Minor,
            Build,
            Revision
        }

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
        /// <summary>
        /// Updates the assembly version
        /// </summary>
        /// <param name="feature">The value for the feature item.</param>
        /// <param name="bugfix">The value for the bug fix item.</param>
        /// <param name="build">The value for the build item.</param>
        private void UpVersion(VersionTypes versionType)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Project selProject = GetSelectedProject();
            string title = "VersionUp V2";
            string message = "Please select the project!";
            Version assemblyVersion;
            string newAssemblyVersion = "";

            if (selProject != null)
            {
                try
                {                   
                    assemblyVersion = new Version(selProject.Properties.Item("AssemblyVersion").Value.ToString());

                    int major = 0, minor = 0, revision = 0, build = 0;

                    switch(versionType)
                    {
                        case VersionTypes.Major:
                            major = assemblyVersion.Major + 1;
                            break;
                        case VersionTypes.Minor:
                            major = assemblyVersion.Major;
                            minor = assemblyVersion.Minor + 1;
                            break;
                        case VersionTypes.Build:
                            major = assemblyVersion.Major;
                            minor = assemblyVersion.Minor;
                            build = assemblyVersion.Build + 1;
                            break;
                        default:
                            major = assemblyVersion.Major;
                            minor = assemblyVersion.Minor;
                            build = assemblyVersion.Build;
                            revision = assemblyVersion.Revision + 1;
                            break;
                    }
                    

                    newAssemblyVersion = $"{major}.{minor}.{build}.{revision}";
                    selProject.Properties.Item("AssemblyVersion").Value = newAssemblyVersion;
                    try
                    {
                        selProject.Properties.Item("AssemblyFileVersion").Value = newAssemblyVersion;
                    }
                    catch
                    {
                        // .NetCore Project doesn't has a property AssemblyFileVersion
                        selProject.Properties.Item("FileVersion").Value = newAssemblyVersion;
                        selProject.Properties.Item("Version").Value = newAssemblyVersion; // Package Version
                    }

                    message = selProject.Name + " " + assemblyVersion.ToString() + " -> " + newAssemblyVersion;
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
        private void MenuMinorCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(VersionTypes.Minor);
        }

        private void MenuMajorCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(VersionTypes.Major);
        }

        private void MenuBuildCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(VersionTypes.Build);
        }

        private void MenuRevisionCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.UpVersion(VersionTypes.Revision);
        }
    }
}
