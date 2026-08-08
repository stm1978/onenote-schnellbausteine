using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    [ComVisible(true)]
    [Guid("CCF3B64A-9580-4D44-BDCE-49BFE95AC510")]
    [ProgId("Schnellbausteine.OneNoteAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private readonly SnippetRepository repository = new SnippetRepository(SnippetRepository.DefaultRoot);
        private readonly StaDispatcher dispatcher = new StaDispatcher();
        private Action<IntPtr> saveSelectionAction;
        private Action<IntPtr> openSnippetMenuAction;
        private int operationActive;

        public AddIn()
        {
            saveSelectionAction = SaveSelectionCore;
            openSnippetMenuAction = OpenSnippetMenuCore;
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInstance, ref Array custom)
        {
            Log.Trace("addin.connection", "entry");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Log.Trace("addin.update", "entry");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log.Trace("addin.startup", "complete");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log.Trace("addin.shutdown", "begin");
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Log.Trace("addin.disconnection", "entry");
            bool dispatcherStopped = dispatcher.Shutdown(TimeSpan.FromSeconds(5));
            if (!dispatcherStopped)
            {
                Log.Info("STA-UI-Thread reagierte beim Beenden nicht.");
            }
            Log.Trace("addin.disconnection", "exit");
            Log.Flush(TimeSpan.FromMilliseconds(500));
        }

        public string GetCustomUI(string ribbonId)
        {
            Log.Trace("ribbon.GetCustomUI", "entry");
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Schnellbausteine.AddIn.Ribbon.xml"))
            using (var reader = new StreamReader(stream))
            {
                string xml = reader.ReadToEnd();
                Log.Trace("ribbon.GetCustomUI", "exit");
                return xml;
            }
        }

        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            Log.Trace("ribbon.loaded", "entry");
        }

        public bool CanSaveSelection(IRibbonControl control)
        {
            Log.Trace("callback.CanSaveSelection", "entry-no-com");
            return true;
        }

        internal bool HasCurrentSelection()
        {
            Log.Trace("diagnostic.HasCurrentSelection", "entry");
            try
            {
                using (var oneNote = new OneNoteClient())
                {
                    bool hasSelection = OneNoteXml.HasSelection(oneNote.ReadCurrentPage());
                    Log.Trace("diagnostic.HasCurrentSelection", "exit");
                    return hasSelection;
                }
            }
            catch (Exception exception)
            {
                Log.Error("Auswahlstatus", exception);
                return false;
            }
        }

        public void SaveSelection(IRibbonControl control)
        {
            Log.Trace("callback.SaveSelection", "entry");
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                DispatchUserAction(
                    "SaveSelection",
                    callbackBodyCompleted.Task,
                    "Der Schnellbaustein konnte nicht gespeichert werden.",
                    saveSelectionAction);
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
                Log.Trace("callback.SaveSelection", "exit");
            }
        }

        private void SaveSelectionCore(IntPtr fallbackOwnerHandle)
        {
            Log.Trace("command.SaveSelection", "entry");
            XDocument fragment;
            IntPtr ownerHandle;
            using (var oneNote = new OneNoteClient())
            {
                Log.Trace("onenote.selection.extract", "before");
                fragment = OneNoteXml.ExtractSelection(oneNote.ReadCurrentPage(), oneNote.ReadBinary);
                Log.Trace("onenote.selection.extract", "after");
                ownerHandle = oneNote.CurrentWindowHandle;
            }

            var owner = new WindowHandleOwner(
                ownerHandle == IntPtr.Zero ? fallbackOwnerHandle : ownerHandle);
            using (var dialog = new NameDialog("Als Schnellbaustein speichern", string.Empty))
            {
                while (true)
                {
                    Log.Trace("ui.NameDialog", "before");
                    DialogResult dialogResult = dialog.ShowDialog(owner);
                    Log.Trace("ui.NameDialog", "after-" + dialogResult);
                    if (dialogResult != DialogResult.OK) break;
                    if (dialog.SnippetName.Length == 0)
                    {
                        Log.Trace("ui.NameRequiredMessage", "before");
                        MessageBox.Show(owner, "Bitte einen Namen eingeben.", "Name erforderlich",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Log.Trace("ui.NameRequiredMessage", "after");
                        continue;
                    }

                    bool overwrite = false;
                    if (repository.ContainsName(dialog.SnippetName, null))
                    {
                        Log.Trace("ui.NameConflictMessage", "before");
                        DialogResult answer = MessageBox.Show(owner,
                            "Ein Schnellbaustein mit diesem Namen existiert bereits.\n\nJa: überschreiben\nNein: anderen Namen eingeben",
                            "Namenskonflikt", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                        Log.Trace("ui.NameConflictMessage", "after-" + answer);
                        if (answer == DialogResult.Cancel) return;
                        if (answer == DialogResult.No) continue;
                        overwrite = true;
                    }

                    Log.Trace("repository.save", "before");
                    repository.Save(dialog.SnippetName, fragment, overwrite);
                    Log.Trace("repository.save", "after");
                    Log.Trace("command.SaveSelection", "exit-saved");
                    return;
                }
            }
            Log.Trace("command.SaveSelection", "exit-cancelled");
        }

        public void OpenSnippetMenu(IRibbonControl control)
        {
            Log.Trace("callback.OpenSnippetMenu", "entry");
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                DispatchUserAction(
                    "OpenSnippetMenu",
                    callbackBodyCompleted.Task,
                    "Die Schnellbausteine konnten nicht geöffnet werden.",
                    openSnippetMenuAction);
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
                Log.Trace("callback.OpenSnippetMenu", "exit");
            }
        }

        private void OpenSnippetMenuCore(IntPtr fallbackOwnerHandle)
        {
            Log.Trace("command.OpenSnippetMenu", "entry");
            IList<Snippet> snippets = repository.GetAll();
            OneNoteClient oneNote = null;
            XDocument anchoredPage = null;
            Exception selectionException = null;
            IntPtr ownerHandle = fallbackOwnerHandle;
            try
            {
                if (snippets.Count > 0)
                {
                    oneNote = new OneNoteClient();
                    try
                    {
                        anchoredPage = oneNote.ReadCurrentPage();
                        OneNoteSelectionInfo selection =
                            OneNoteXml.AnalyzeSelection(anchoredPage);
                        Log.Diagnostic(
                            "onenote.selection.capture",
                            selection.ToDiagnosticString());
                        if (!selection.IsUsable)
                        {
                            selectionException =
                                new InvalidOperationException(selection.Problem);
                        }
                    }
                    catch (Exception exception)
                    {
                        selectionException = exception;
                    }
                    if (oneNote.CurrentWindowHandle != IntPtr.Zero)
                    {
                        ownerHandle = oneNote.CurrentWindowHandle;
                    }
                }

                var owner = new WindowHandleOwner(
                    ownerHandle == IntPtr.Zero ? fallbackOwnerHandle : ownerHandle);
                SnippetMenuChoice choice;
                using (ISnippetMenuPopup popup = new SnippetMenuPopup(snippets))
                {
                    choice = popup.Choose(owner);
                }
                Log.Trace("ui.SnippetMenuPopup", "disposed");

                SnippetMenuChoiceDispatcher.Execute(
                    choice,
                    delegate(Snippet snippet)
                    {
                        if (selectionException != null) throw selectionException;
                        if (oneNote == null || anchoredPage == null)
                        {
                            throw new InvalidOperationException(
                                "Die OneNote-Auswahl wurde nicht verankert.");
                        }
                        InsertSnippetOperation.Execute(
                            oneNote,
                            anchoredPage,
                            snippet,
                            repository.LoadFragment);
                    },
                    delegate
                    {
                        if (oneNote != null)
                        {
                            oneNote.Dispose();
                            oneNote = null;
                        }
                        ShowManageDialog(owner);
                    });
            }
            finally
            {
                if (oneNote != null) oneNote.Dispose();
            }
            Log.Trace("command.OpenSnippetMenu", "exit");
        }

        private void ShowManageDialog(IWin32Window owner)
        {
            using (var dialog = new ManageDialog(repository, delegate { }))
            {
                Log.Trace("ui.ManageDialog", "before");
                dialog.ShowDialog(owner);
                Log.Trace("ui.ManageDialog", "after");
            }
        }

        private void DispatchUserAction(
            string operation, Task callbackBodyCompleted, string errorMessage, Action<IntPtr> action)
        {
            Log.Trace("dispatch." + operation, "entry");
            if (Interlocked.CompareExchange(ref operationActive, 1, 0) != 0)
            {
                Log.Trace("dispatch." + operation, "rejected-busy");
                return;
            }

            IntPtr ownerHandle = NativeMethods.GetForegroundWindow();
            try
            {
                dispatcher.PostAfterCallback(
                    callbackBodyCompleted,
                    delegate
                    {
                        Log.Trace("dispatch." + operation, "action.entry");
                        try
                        {
                            action(ownerHandle);
                        }
                        catch (Exception exception)
                        {
                            ShowError(errorMessage, exception, ownerHandle);
                        }
                        finally
                        {
                            Interlocked.Exchange(ref operationActive, 0);
                            Log.Trace("dispatch." + operation, "action.exit");
                        }
                    },
                    delegate(Exception exception)
                    {
                        Interlocked.Exchange(ref operationActive, 0);
                        Log.Error(errorMessage, exception);
                    });
            }
            catch (Exception exception)
            {
                Interlocked.Exchange(ref operationActive, 0);
                Log.Error(errorMessage, exception);
            }
            Log.Trace("dispatch." + operation, "exit");
        }

        private static void ShowError(string message, Exception exception, IntPtr ownerHandle)
        {
            Log.Error(message, exception);
            Log.Trace("ui.ErrorMessage", "before");
            MessageBox.Show(new WindowHandleOwner(ownerHandle), message + "\n\n" + exception.Message,
                "Schnellbausteine", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Log.Trace("ui.ErrorMessage", "after");
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll")]
            internal static extern IntPtr GetForegroundWindow();
        }
    }
}
