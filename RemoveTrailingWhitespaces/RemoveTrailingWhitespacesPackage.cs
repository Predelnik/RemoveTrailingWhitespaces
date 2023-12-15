using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Editor;
using System.Linq;
using EnvDTE;
using System.Collections.Generic;

using Task = System.Threading.Tasks.Task;
using Microsoft;

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
        readonly RemoveTrailingWhitespacesPackage _pkg;

        public RunningDocTableEvents(RemoveTrailingWhitespacesPackage pkg)
        {
            _pkg = pkg;
        }

        public int OnBeforeSave(uint docCookie)
        {
            if (_pkg.RemoveOnSave())
            {
                _pkg.RemoveTrailingWhiteSpaces(docCookie);
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
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidRemoveTrailingWhitespacesPkgString)]
    [ProvideOptionPage(typeof(OptionsPage), "Remove Trailing Whitespaces", "Options", 1000, 1001, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}", PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideUIContextRule("{f1536ef8-92ec-443c-9ed7-fdadf150da82}",
        name: "Trigger for autoloading the RemoveTrailingWhitespaces extension",
        expression: "DocOpen",
        termNames: new[] { "DocOpen" },
        termValues: new[] { "HierSingleSelectionName:.$" })]
    public sealed class RemoveTrailingWhitespacesPackage : AsyncPackage
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
        public _DTE dte;
        public IVsRunningDocumentTable rdt;
        public IFindService findService;
        private uint rdtCookie;
        public IComponentModel componentModel;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override async Task InitializeAsync(System.Threading.CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            dte = await GetServiceAsync(typeof(_DTE)) as _DTE;
            Assumes.Present(dte);
            rdt = await GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            Assumes.Present(rdt);
            componentModel = GetGlobalService(typeof(SComponentModel)) as IComponentModel;
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            InitializePackage();
        }

        private void InitializePackage ()
        {
            rdt.AdviseRunningDocTableEvents(new RunningDocTableEvents(this), out rdtCookie);
            if (GetService(typeof(IMenuCommandService)) is OleMenuCommandService mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(
                    GuidList.guidRemoveTrailingWhitespacesCmdSet, (int)PkgCmdIDList.cmdIdRemoveTrailingWhitespaces);
                OleMenuCommand menuItem = new OleMenuCommand(OnRemoveTrailingWhitespacesPressed, menuCommandID);
                menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
                mcs.AddCommand(menuItem);
            }
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;

            cmd.Visible = IsNeededForActiveDocument();
            cmd.Enabled = cmd.Visible;
        }

        private bool IsNeededForActiveDocument()
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

            if (!(doc.Object("TextDocument") is TextDocument))
            {
                return false;
            }

            return true;
        }

        private void OnRemoveTrailingWhitespacesPressed(object sender, EventArgs e)
        {
            if (dte.ActiveDocument == null) return;
            if (!(dte.ActiveDocument.Object() is TextDocument textDocument)) return;

            uint docCookie = GetDocCookie(dte.ActiveDocument.FullName);
            RemoveTrailingWhiteSpaces(docCookie);
        }

        private IFinder GetFinder(string findWhat, string replacement, ITextBuffer textBuffer)
        {
            var findService = componentModel.GetService<IFindService> ();
            var finderFactory = findService.CreateFinderFactory(findWhat, replacement, FindOptions.UseRegularExpressions);
            return finderFactory.Create(textBuffer.CurrentSnapshot);
        }

        internal static ITextBuffer GettextBufferAt(IVsTextBuffer textBuffer, IComponentModel componentModel)
        {
            var editorAdapterFactoryService = componentModel.GetService<IVsEditorAdaptersFactoryService>();
            return editorAdapterFactoryService.GetDataBuffer(textBuffer);
        }

        private static void ReplaceAll(ITextBuffer textBuffer, IEnumerable<FinderReplacement> replacements)
        {
            if (replacements.Any())
            {
                using (var edit = textBuffer.CreateEdit())
                {
                    foreach (var match in replacements)
                    {
                        edit.Replace(match.Match, match.Replace);
                    }

                    edit.Apply();
                }
            }
        }

        public uint GetDocCookie(string docFullName)
        {
            IVsHierarchy hierarchy = null;
            uint itemid = 0;
            IntPtr docDataUnk = IntPtr.Zero;
            uint lockCookie = 0;

            IEnumRunningDocuments allDocs;
            if (VSConstants.S_OK != rdt.GetRunningDocumentsEnum(out allDocs))
                return 0;
            uint[] array = new uint[1];
            uint pceltFetched = 0;
            while (VSConstants.S_OK == allDocs.Next(1, array, out pceltFetched) && (pceltFetched == 1))
            {
                uint pgrfRDTFlags;
                uint pdwReadLocks;
                uint pdwEditLocks;
                string pbstrMkDocument;
                IVsHierarchy ppHier;
                uint pitemid;
                IntPtr ppunkDocData;
                rdt.GetDocumentInfo(array[0], out pgrfRDTFlags, out pdwReadLocks, out pdwEditLocks, out pbstrMkDocument, out ppHier, out pitemid, out ppunkDocData);
                if (pbstrMkDocument == docFullName)
                    return array[0];
            }

            return 0;
        }

        public void RemoveTrailingWhiteSpaces(uint docCookie)
        {
            RunningDocumentInfo runningDocumentInfo = new RunningDocumentInfo(rdt, docCookie);

            IVsHierarchy hierarchy = null;
            uint itemid = 0;
            IntPtr docDataUnk = IntPtr.Zero;
            uint lockCookie = 0;

            int hr = rdt.FindAndLockDocument((uint)_VSRDTFLAGS.RDT_ReadLock, runningDocumentInfo.Moniker, out hierarchy, out itemid, out docDataUnk, out lockCookie);
            if (hr != VSConstants.S_OK || !(Marshal.GetUniqueObjectForIUnknown(docDataUnk) is IVsTextBuffer vsTextBuffer))
                return;

            var textBuffer = GettextBufferAt(vsTextBuffer, componentModel);
            ReplaceAll(textBuffer, GetFinder("[^\\S\\r\\n]+(?=\\r?$)", "", textBuffer).FindForReplaceAll());
        }

        public bool RemoveOnSave()
        {
            var props = dte.get_Properties("Remove Trailing Whitespaces", "Options");
            return (bool)props.Item("RemoveTrailingWhitespacesOnSave").Value;
        }


        #endregion

    }
}
