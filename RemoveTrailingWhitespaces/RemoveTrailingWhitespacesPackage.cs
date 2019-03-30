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
using System.Linq;
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

    internal class RunningDocTableEvents : IVsRunningDocTableEvents3
    {
        RemoveTrailingWhitespacesPackage _pkg;

        public RunningDocTableEvents(RemoveTrailingWhitespacesPackage pkg)
        {
            _pkg = pkg;
        }

        public int OnBeforeSave(uint docCookie)
        {
            if (_pkg.removeOnSave())
            {
                RunningDocumentInfo runningDocumentInfo = _pkg.rdt.GetDocumentInfo(docCookie);
                EnvDTE.Document document = _pkg.dte.Documents.OfType<EnvDTE.Document>().SingleOrDefault(x => x.FullName == runningDocumentInfo.Moniker);
                if (document == null)
                    return VSConstants.S_OK;
                var textDoc = document.Object("TextDocument") as TextDocument;
                if (textDoc != null)
                    RemoveTrailingWhitespacesPackage.removeTrailingWhiteSpaces(textDoc);
            }
            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs) { return VSConstants.S_OK; }
        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld,
                                            uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew,
                                            uint itemidNew, string pszMkDocumentNew)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame) { return VSConstants.S_OK; }
        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie) { return VSConstants.S_OK; }
        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame) { return VSConstants.S_OK; }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }
    }

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
        public DTE dte;
        private Properties _props;
        public RunningDocumentTable rdt;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();
            dte = GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
            _props = dte.get_Properties("Remove Trailing Whitespaces", "Options");
            rdt = new RunningDocumentTable(this);
            rdt.Advise(new RunningDocTableEvents(this));
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
            var doc = dte.ActiveDocument;
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
            if (dte.ActiveDocument == null) return;
            var textDocument = dte.ActiveDocument.Object() as TextDocument;
            if (textDocument == null) return;
            removeTrailingWhiteSpaces(textDocument);
        }

        public static void removeTrailingWhiteSpaces(TextDocument textDocument)
        {
            textDocument.ReplacePattern("[^\\S\\r\\n]+(?=\\r?$)", "", (int)vsFindOptions.vsFindOptionsRegularExpression);
        }

        public bool removeOnSave()
        {
            return (bool)_props.Item("RemoveTrailingWhitespacesOnSave").Value;
        }


        #endregion

    }
}
