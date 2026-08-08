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
        private readonly SnippetRepository repository;
        private readonly SnippetGalleryModel gallery;
        private readonly StaDispatcher dispatcher = new StaDispatcher();
        private Action<IntPtr> saveSelectionAction;
        private Action<Snippet, IntPtr> insertSnippetAction;
        private Action<IntPtr> manageSnippetsAction;
        private IRibbonUI ribbonUi;
        private int operationActive;

        public AddIn()
            : this(new SnippetRepository(SnippetRepository.DefaultRoot))
        {
        }

        internal AddIn(SnippetRepository repository)
        {
            if (repository == null) throw new ArgumentNullException("repository");
            this.repository = repository;
            gallery = new SnippetGalleryModel(repository);
            saveSelectionAction = SaveSelectionCore;
            insertSnippetAction = InsertSnippetCore;
            manageSnippetsAction = ManageSnippetsCore;
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
            this.ribbonUi = ribbonUi;
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
                    InvalidateSnippetGallery();
                    Log.Trace("command.SaveSelection", "exit-saved");
                    return;
                }
            }
            Log.Trace("command.SaveSelection", "exit-cancelled");
        }

        public int GetSnippetCount(IRibbonControl control)
        {
            Log.Trace("callback.GetSnippetCount", "entry-no-com");
            return gallery.RefreshAndGetCount();
        }

        public string GetSnippetLabel(IRibbonControl control, int itemIndex)
        {
            Log.Trace("callback.GetSnippetLabel", "entry-no-com");
            return gallery.GetLabel(itemIndex);
        }

        public string GetSnippetId(IRibbonControl control, int itemIndex)
        {
            Log.Trace("callback.GetSnippetId", "entry-no-com");
            return gallery.GetId(itemIndex);
        }

        public bool GetSnippetGalleryEnabled(IRibbonControl control)
        {
            Log.Trace("callback.GetSnippetGalleryEnabled", "entry-no-com");
            return gallery.RefreshAndGetEnabled();
        }

        public void InsertGallerySnippet(
            IRibbonControl control,
            string selectedItemId,
            int selectedItemIndex)
        {
            Log.Trace("callback.InsertGallerySnippet", "entry");
            Snippet snippet = null;
            Exception itemException = null;
            try
            {
                snippet = gallery.GetSnippet(selectedItemId, selectedItemIndex);
            }
            catch (Exception exception)
            {
                itemException = exception;
            }
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                DispatchUserAction(
                    "InsertGallerySnippet",
                    callbackBodyCompleted.Task,
                    "Der Schnellbaustein konnte nicht eingefügt werden.",
                    delegate(IntPtr ownerHandle)
                    {
                        if (itemException != null) throw itemException;
                        insertSnippetAction(snippet, ownerHandle);
                    });
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
                Log.Trace("callback.InsertGallerySnippet", "exit");
            }
        }

        private void InsertSnippetCore(Snippet snippet, IntPtr fallbackOwnerHandle)
        {
            Log.Trace("command.InsertGallerySnippet", "entry");
            using (var oneNote = new OneNoteClient())
            {
                Log.Trace("onenote.gallery-selection.read.piBinaryDataSelection", "before");
                XDocument anchoredPage = oneNote.ReadCurrentPage();
                Log.Trace("onenote.gallery-selection.read.piBinaryDataSelection", "after");
                OneNoteSelectionInfo selection = OneNoteXml.AnalyzeSelection(anchoredPage);
                Log.Diagnostic(
                    "onenote.gallery-selection.after-item-action",
                    selection.ToDiagnosticString());
                if (!selection.IsUsable)
                {
                    Log.Trace("command.InsertGallerySnippet", "abort-selection");
                    throw new InvalidOperationException(selection.Problem);
                }

                InsertSnippetOperation.Execute(
                    oneNote,
                    anchoredPage,
                    snippet,
                    repository.LoadFragment);
            }
            Log.Trace("command.InsertGallerySnippet", "exit");
        }

        public void ManageSnippets(IRibbonControl control)
        {
            Log.Trace("callback.ManageSnippets", "entry");
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                DispatchUserAction(
                    "ManageSnippets",
                    callbackBodyCompleted.Task,
                    "Die Schnellbausteine-Verwaltung konnte nicht geöffnet werden.",
                    manageSnippetsAction);
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
                Log.Trace("callback.ManageSnippets", "exit");
            }
        }

        private void ManageSnippetsCore(IntPtr fallbackOwnerHandle)
        {
            var owner = new WindowHandleOwner(fallbackOwnerHandle);
            using (var dialog = new ManageDialog(repository, InvalidateSnippetGallery))
            {
                Log.Trace("ui.ManageDialog", "before");
                dialog.ShowDialog(owner);
                Log.Trace("ui.ManageDialog", "after");
            }
        }

        private void InvalidateSnippetGallery()
        {
            gallery.Invalidate();
            IRibbonUI currentRibbonUi = ribbonUi;
            if (currentRibbonUi == null) return;
            Log.Trace("ribbon.invalidate.sqGallery", "before");
            currentRibbonUi.InvalidateControl("sqGallery");
            Log.Trace("ribbon.invalidate.sqGallery", "after");
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

    internal sealed class SnippetGalleryModel
    {
        private const string EmptyLabel = "Keine Schnellbausteine gespeichert";
        private readonly SnippetRepository repository;
        private readonly object gate = new object();
        private IList<Snippet> snapshot = new List<Snippet>();
        private bool valid;

        public SnippetGalleryModel(SnippetRepository repository)
        {
            if (repository == null) throw new ArgumentNullException("repository");
            this.repository = repository;
        }

        public int RefreshAndGetCount()
        {
            lock (gate)
            {
                Refresh();
                return Math.Max(1, snapshot.Count);
            }
        }

        public bool RefreshAndGetEnabled()
        {
            lock (gate)
            {
                Refresh();
                return snapshot.Count > 0;
            }
        }

        public string GetLabel(int itemIndex)
        {
            lock (gate)
            {
                EnsureSnapshot();
                return snapshot.Count == 0
                    ? GetEmptyItem(itemIndex, EmptyLabel)
                    : GetItem(itemIndex).Name;
            }
        }

        public string GetId(int itemIndex)
        {
            lock (gate)
            {
                EnsureSnapshot();
                return snapshot.Count == 0
                    ? GetEmptyItem(itemIndex, "sqEmpty")
                    : "sqSnippet_" + GetItem(itemIndex).Id;
            }
        }

        public Snippet GetSnippet(string itemId, int itemIndex)
        {
            lock (gate)
            {
                EnsureSnapshot();
                if (snapshot.Count == 0)
                {
                    throw new InvalidOperationException(
                        "Es sind keine Schnellbausteine zum Einfügen gespeichert.");
                }
                Snippet item = GetItem(itemIndex);
                if (!string.Equals(
                    itemId,
                    "sqSnippet_" + item.Id,
                    StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Die ausgewählte Galerieposition ist nicht mehr eindeutig.");
                }
                return item;
            }
        }

        public void Invalidate()
        {
            lock (gate)
            {
                valid = false;
            }
        }

        private void EnsureSnapshot()
        {
            if (!valid)
            {
                throw new InvalidOperationException(
                    "Die Schnellbaustein-Galerie wurde noch nicht aktualisiert.");
            }
        }

        private void Refresh()
        {
            snapshot = repository.GetAll()
                .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            valid = true;
        }

        private Snippet GetItem(int itemIndex)
        {
            if (itemIndex < 0 || itemIndex >= snapshot.Count)
            {
                throw new ArgumentOutOfRangeException("itemIndex");
            }
            return snapshot[itemIndex];
        }

        private static T GetEmptyItem<T>(int itemIndex, T value)
        {
            if (itemIndex != 0) throw new ArgumentOutOfRangeException("itemIndex");
            return value;
        }
    }
}
