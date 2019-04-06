using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft;

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
        public const int CommandFeature = 0x0100;
        public const int CommandBugFix = 0x0200;
        public const int CommandBuild = 0x0300;

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

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandFeature = new CommandID(CommandSet, CommandFeature);
                var menuFeature = new MenuCommand(this.MenuFeatureCallback, menuCommandFeature);
                commandService.AddCommand(menuFeature);

                var menuCommandBugFix = new CommandID(CommandSet, CommandBugFix);
                var menuBugFix = new MenuCommand(this.MenuBugFixCallback, menuCommandBugFix);
                commandService.AddCommand(menuBugFix);

                var menuCommandBuild = new CommandID(CommandSet, CommandBuild);
                var menuBuild = new MenuCommand(this.MenuBuildCallback, menuCommandBuild);
                commandService.AddCommand(menuBuild);
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
        public static void Initialize(Package package)
        {
            Instance = new ContextMenu(package);
        }

        private string GetActiveFilePath(IServiceProvider serviceProvider)
        {
            EnvDTE80.DTE2 applicationObject = serviceProvider.GetService(typeof(DTE)) as EnvDTE80.DTE2;
            return applicationObject.ActiveDocument.FullName;
        }
        /// <summary>
        /// Updates the assembly version
        /// </summary>
        /// <param name="feature">The value for the feature item.</param>
        /// <param name="bugfix">The value for the bugfix item.</param>
        /// <param name="build">The value for the build item.</param>
        private void UpVersion(int feature = 0, int bugfix = 0, int build = 0)
        {
            Project selProject = GetSelectedProject();
            string title = "VersionUp V2";
            string message = "Please select the project!";
            Version assemblyVersion;
            string newAssemblyVersion = "";

            if (selProject != null)
            {
                int count = selProject.Properties.Count;
                try
                {                   
                    assemblyVersion = new Version(selProject.Properties.Item("AssemblyVersion").Value.ToString());
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


                    int buildItem = assemblyVersion.Revision + build;
                    int bugfixItem = assemblyVersion.Build + bugfix;
                    int featureItem = assemblyVersion.Minor + feature;
 
                    newAssemblyVersion = $"{assemblyVersion.Major}.{featureItem}.{bugfixItem}.{buildItem}";
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
        private void MenuFeatureCallback(object sender, EventArgs e)
        {
            this.UpVersion(feature: 1);
        }

        private void MenuBugFixCallback(object sender, EventArgs e)
        {
            this.UpVersion(bugfix: 1);
        }

        private void MenuBuildCallback(object sender, EventArgs e)
        {
            this.UpVersion(build: 1);
        }
    }
}
