// ***********************************************************************
// Assembly         : FindByWC3
// Author           : Prasanna
// Created          : 04-09-2019
//
// Last Modified By : Prasanna
// Last Modified On : 04-10-2019
// ***********************************************************************
// <copyright file="FindByWCcommand.cs" company="Prasanna">
//     Copyright (c) Prasanna 2019 . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Reflection;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TeamFoundation.VersionControl;
using Task = System.Threading.Tasks.Task;


namespace FindByWC3
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class FindByWCcommand
    {
        /// <summary>
        /// The dte3
        /// </summary>
        private DTE2 dte3;

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a42f6f18-5942-4261-b06b-1ea60f9ecf6e");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="FindByWCcommand" /> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        /// <exception cref="ArgumentNullException">
        /// package
        /// or
        /// commandService
        /// </exception>
        private FindByWCcommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            IVsExtensibility extensibility =
                (IVsExtensibility)package.GetServiceAsync(typeof(IVsExtensibility)).Result;

            dte3 = extensibility.GetGlobalsObject(null).DTE as DTE2;

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        /// <value>The instance.</value>
        public static FindByWCcommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        /// <value>The service provider.</value>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        /// <returns>Task.</returns>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in FindByWCcommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new FindByWCcommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        /// <exception cref="InvalidOperationException">
        /// Something is wrong! Please Contact your Administrator
        /// or
        /// Something is wrong! Please Contact your Administrator
        /// </exception>
        private void Execute(object sender, EventArgs e)
        {
            string title = "Find By WildCard - Visual Studio 2017";
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                //Followed by this to get the Version
                if (!(dte3.GetObject("Microsoft.VisualStudio.TeamFoundation.VersionControl.VersionControlExt") is VersionControlExt versionControlExt))
                {
                    return;
                }

                var versionControlExplorerExt = versionControlExt.Explorer;

                string path = versionControlExplorerExt.CurrentFolderItem.SourceServerPath;

                string url = versionControlExplorerExt.Workspace.VersionControlServer.TeamProjectCollection.Uri.OriginalString;

                if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(url))
                {
                    VsShellUtilities.ShowMessageBox(
                        package,
                        @"Could not get TFS Selected path or URL! Manually provide to Search!",
                        title,
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }

                string assemblyLoc = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string[] files =
                    Directory.GetFiles(
                        assemblyLoc ??
                        throw new InvalidOperationException("Something is wrong! Please Contact your Administrator"),
                        "*.exe", SearchOption.AllDirectories);

                string fileName = files.FirstOrDefault();

                using (System.Diagnostics.Process myProcess = new System.Diagnostics.Process())
                {
                    myProcess.StartInfo.UseShellExecute = false;
                    myProcess.StartInfo.FileName = fileName ?? throw new InvalidOperationException("Something is wrong! Please Contact your Administrator");
                    myProcess.StartInfo.CreateNoWindow = true;
                    myProcess.StartInfo.Arguments = $"{url} {path}";
                    myProcess.Start();
                }
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    package,
                    ex.ToString(),
                    title,
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }
    }
}
