//------------------------------------------------------------------------------
// <copyright file="LearnToolWindowPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace LearnVS
{
    // The minimum requirement for a class to be considered a valid package for Visual Studio
    // is to implement the IVsPackage interface and register itself with the shell.
    // This package uses the helper classes defined inside the Managed Package Framework (MPF)
    // to do it: it derives from the Package class that provides the implementation of the
    // IVsPackage interface and uses the registration attributes defined in the framework to
    // register itself and its components with the shell. These attributes tell the pkgdef creation
    // utility what data to put into .pkgdef file.
    // To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(LearnToolWindow), Style=VsDockStyle.Tabbed, Orientation=ToolWindowOrientation.Right)]
    [Guid(LearnToolWindowPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class LearnToolWindowPackage : Package
    {
        /// <summary>
        /// LearnToolWindowPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "423c0e10-e2b1-4123-a47b-bd7367834869";

        // Cache the Menu Command Service since we will use it multiple times
        private OleMenuCommandService menuService;

        /// <summary>
        /// Initializes a new instance of the <see cref="LearnToolWindow"/> class.
        /// </summary>
        public LearnToolWindowPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        internal OleMenuCommand DefineCommandHandler(EventHandler handler, CommandID id)
        {
            // if the package is zombied, we don't want to add commands
            if (Zombied)
                return null;

            // Make sure we have the service
            if (menuService == null)
            {
                // Get the OleCommandService object provided by the MPF; this object is the one
                // responsible for handling the collection of commands implemented by the package.
                menuService = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            }
            OleMenuCommand command = null;
            if (null != menuService)
            {
                // Add the command handler
                command = new OleMenuCommand(handler, id);
                menuService.AddCommand(command);
            }
            return command;
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            LearnToolWindowCommand.Initialize(this);
            base.Initialize();
        }

        #endregion
    }
}
