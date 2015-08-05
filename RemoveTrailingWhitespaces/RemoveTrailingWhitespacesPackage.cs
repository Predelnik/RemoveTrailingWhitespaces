using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using EnvDTE;

namespace Predelnik.RemoveTrailingWhitespaces
{
    [CLSCompliant(false), ComVisible(true)]
    public class OptionsPage : DialogPage
    {
        private bool removeTrailingWhitespacesOnSave = true;
        [Category("All")]
        [DisplayName("Remove Trailing Whitespaces on Save")]
        public bool RemoveTrailingWhitespacesOnSave
        {
            get { return removeTrailingWhitespacesOnSave; }
            set { removeTrailingWhitespacesOnSave = value; }
        }
    };

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidRemoveTrailingWhitespacesPkgString)]
    [ProvideOptionPage(typeof(OptionsPage), "Remove Trailing Whitespaces", "Options", 1000, 1001, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}")]
    public sealed class RemoveTrailingWhitespacesPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require
        /// any Visual Studio service because at this point the package object is created but
        /// not sited yet inside Visual Studio environment. The place to do all the other
        /// initialization is the Initialize method.
        /// </summary>

        public RemoveTrailingWhitespacesPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members
        private DTE _dte;
        private Properties _props;
        DocumentEvents _docEvents;
        Events _events;
        bool _actionAppliedFlag = false;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();
            _dte = GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
            _props = _dte.get_Properties("Remove Trailing Whitespaces", "Options");
            _events = _dte.Events;
            _docEvents = _events.DocumentEvents;
            _docEvents.DocumentSaved += onDocumentSaved;
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs != null)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(
                    GuidList.guidRemoveTrailingWhitespacesCmdSet, (int)PkgCmdIDList.cmdIdRemoveTrailingWhitespaces);
                OleMenuCommand menuItem = new OleMenuCommand(onRemoveTrailingWhitespacesPressed, menuCommandID);
                menuItem.BeforeQueryStatus += onBeforeQueryStatus;
                mcs.AddCommand(menuItem);
            }
        }

        private void onBeforeQueryStatus(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;

            cmd.Visible = isNeededForActiveDocument();
            cmd.Enabled = cmd.Visible;
        }

        private bool isNeededForActiveDocument()
        {
            var doc = _dte.ActiveDocument;
            if (doc == null)
            {
                return false;
            }

            if (doc.ReadOnly)
            {
                return false;
            }

            var textDoc = doc.Object("TextDocument") as TextDocument;
            if (textDoc == null)
            {
                return false;
            }

            return true;
        }

        private void onRemoveTrailingWhitespacesPressed(object sender, EventArgs e)
        {
            if (_dte.ActiveDocument == null) return;
            var textDocument = _dte.ActiveDocument.Object() as TextDocument;
            if (textDocument == null) return;
            removeTrailingWhiteSpaces(textDocument);
        }

        private void onDocumentSaved(Document Document)
        {
            if (Document == null || !removeOnSave())
                return;
            var textDocument = Document.Object() as TextDocument;
            if (textDocument == null) return;

            if (!_actionAppliedFlag)
            {
                try
                {
                    removeTrailingWhiteSpaces(textDocument);
                    _actionAppliedFlag = true;
                    Document.Save();
                }
                catch (Exception ex)
                {
                    Debug.Print("Trailing Whitespace Removal Exception: " + ex.Message);
                }
            }
            else
                _actionAppliedFlag = false;

        }

        private static void removeTrailingWhiteSpaces(TextDocument textDocument)
        {
            textDocument.ReplacePattern("[^\\S\\r\\n]+(?=\\r?$)", "", (int)vsFindOptions.vsFindOptionsRegularExpression);
        }

        private bool removeOnSave()
        {
            return (bool)_props.Item("RemoveTrailingWhitespacesOnSave").Value;
        }


        #endregion

    }
}
