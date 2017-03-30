using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace LearnVS
{
    // Tool windows are composed of a frame (implemented by the shell) and a pane,
    // usually implemented by the package implementer.
    [Guid("c2e58137-32e8-47c0-94c4-eee5c4312081")]
    public class LearnToolWindow : ToolWindowPane
    {
        private LearnToolWindowControl learnToolWindowControl;

        public LearnToolWindow() : base(null)
        {
            this.Caption = "Learn Visual Studio";

            // To get "real" VS toolbars, we're using this Win32 approach, rather than the more user-friendly WPF
            // ToolBar control. Painful, but code that's easy to clone for another project. 
            this.ToolBar = new CommandID(GuidsList.guidClientCmdSet, PkgCmdIds.IDM_MyToolbar);

            // Specify that we want the toolbar at the top of the window (duh!)
            ToolBarLocation = (int)VSTWT_LOCATION.VSTWT_TOP;

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            learnToolWindowControl = new LearnToolWindowControl();

            this.Content = learnToolWindowControl;
        }

        public override void OnToolWindowCreated()
        {
            base.OnToolWindowCreated();

            CommandID idCS = new CommandID(GuidsList.guidClientCmdSet, PkgCmdIds.cmdIdCSharpTutorialButton);
            OleMenuCommand commandCS = DefineCommandHandler(new EventHandler(this.CSharpButtonClicked), idCS);

            CommandID idVB = new CommandID(GuidsList.guidClientCmdSet, PkgCmdIds.cmdIdVBTutorialButton);
            OleMenuCommand commandVB = DefineCommandHandler(new EventHandler(this.VBButtonClicked), idVB);

            CommandID idFS = new CommandID(GuidsList.guidClientCmdSet, PkgCmdIds.cmdIdFSharpTutorialButton);
            OleMenuCommand commandFS = DefineCommandHandler(new EventHandler(this.FSharpButtonClicked), idFS);

        }

        private void CSharpButtonClicked(object sender, EventArgs args)
        {
            learnToolWindowControl.BrowserUri = LearnUris.CSharpTutorial;
        }

        private void VBButtonClicked(object sender, EventArgs args)
        {
            learnToolWindowControl.BrowserUri = LearnUris.VBTutorial;
        }

        private void FSharpButtonClicked(object sender, EventArgs args)
        {
            learnToolWindowControl.BrowserUri = LearnUris.FSharpTutorial;
        }

        private OleMenuCommand DefineCommandHandler(EventHandler handler, CommandID id)
        {
            // First add it to the package. This is to keep the visibility
            // of the command on the toolbar constant when the tool window does
            // not have focus. In addition, it creates the command object for us.
            LearnToolWindowPackage package = (LearnToolWindowPackage)this.Package;
            OleMenuCommand command = package.DefineCommandHandler(handler, id);

            // Verify that the command was added
            if (command == null)
                return command;

            // Get the OleCommandService object provided by the base window pane class; this object is the one
            // responsible for handling the collection of commands implemented by the package.

            if (GetService(typeof(IMenuCommandService)) is OleMenuCommandService menuService)
            {
                // Add the command handler
                menuService.AddCommand(command);
            }
            return command;
        }
    }
}
