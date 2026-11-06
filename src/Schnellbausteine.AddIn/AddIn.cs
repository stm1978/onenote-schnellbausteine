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
        private readonly SnippetMenuModel menu;
        private static readonly string[] SnippetMenuControlIds =
        {
            "sqMain",
            "sqInsertText",
            "sqInsertOutline",
            "sqInsertInk",
            "sqInsertLasso"
        };
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
            menu = new SnippetMenuModel(repository);
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
            Log.Trace("addin.startup", "entry");
            Log.Trace("addin.startup", "exit");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log.Trace("addin.shutdown", "begin");
            try
            {
                if (ribbonUi != null && Marshal.IsComObject(ribbonUi))
                {
                    Log.Trace("ribbon.earlyrelease", "before");
                    Marshal.ReleaseComObject(ribbonUi);
                    ribbonUi = null;
                    Log.Trace("ribbon.earlyrelease", "after");
                }
            }
            catch (Exception ex)
            {
                Log.Error("Fehler beim frühen Freigeben der Ribbon-UI.", ex);
            }
            Log.Trace("addin.shutdown", "end");
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Log.Trace("addin.disconnection", "entry");
            
            bool dispatcherStopped = dispatcher.Shutdown(TimeSpan.FromSeconds(5));
            if (!dispatcherStopped)
            {
                Log.Info("STA-UI-Thread reagierte beim Beenden nicht innerhalb von 5 Sekunden.");
            }
            
            try
            {
                if (ribbonUi != null)
                {
                    Log.Trace("ribbon.cleanup", "before");
                    if (Marshal.IsComObject(ribbonUi))
                    {
                        Marshal.FinalReleaseComObject(ribbonUi);
                    }
                    ribbonUi = null;
                    Log.Trace("ribbon.cleanup", "after");
                }
            }
            catch (Exception ex)
            {
                Log.Error("Fehler beim Freigeben der Ribbon-UI.", ex);
            }
            
            try
            {
                Log.Trace("cleanup.gc", "before");
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                Log.Trace("cleanup.gc", "after");
            }
            catch (Exception ex)
            {
                Log.Error("Fehler beim Garbage Collection.", ex);
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
                    InvalidateSnippetMenu();
                    Log.Trace("command.SaveSelection", "exit-saved");
                    return;
                }
            }
            Log.Trace("command.SaveSelection", "exit-cancelled");
        }

        public string GetSnippetMenuContent(IRibbonControl control)
        {
            Log.Trace("callback.GetSnippetMenuContent", "entry-no-com");
            return menu.RefreshAndGetContent(control == null ? null : control.Id);
        }

        public void InsertSnippet(IRibbonControl control)
        {
            Log.Trace("callback.InsertSnippet", "entry");
            string controlId = control == null ? null : control.Id;
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                DispatchUserAction(
                    "InsertSnippet",
                    callbackBodyCompleted.Task,
                    "Der Schnellbaustein konnte nicht eingefügt werden.",
                    delegate(IntPtr ownerHandle)
                    {
                        insertSnippetAction(menu.GetSnippet(controlId), ownerHandle);
                    });
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
                Log.Trace("callback.InsertSnippet", "exit");
            }
        }

        private void InsertSnippetCore(Snippet snippet, IntPtr fallbackOwnerHandle)
        {
            Log.Trace("command.InsertSnippet", "entry");
            using (var oneNote = new OneNoteClient())
            {
                Log.Trace("onenote.menu-selection.read.piBinaryDataSelection", "before");
                XDocument anchoredPage = oneNote.ReadCurrentPage();
                Log.Trace("onenote.menu-selection.read.piBinaryDataSelection", "after");
                OneNoteSelectionInfo selection = OneNoteXml.AnalyzeSelection(anchoredPage);
                Log.Diagnostic(
                    "onenote.menu-selection.after-item-action",
                    selection.ToDiagnosticString());
                if (!selection.IsUsable)
                {
                    Log.Trace("command.InsertSnippet", "abort-selection");
                    throw new InvalidOperationException(selection.Problem);
                }

                InsertSnippetOperation.Execute(
                    oneNote,
                    anchoredPage,
                    snippet,
                    repository.LoadFragment);
            }
            Log.Trace("command.InsertSnippet", "exit");
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
            using (var dialog = new ManageDialog(repository))
            {
                Log.Trace("ui.ManageDialog", "before");
                dialog.ShowDialog(owner);
                Log.Trace("ui.ManageDialog", "after");
                if (dialog.Changed)
                {
                    try
                    {
                        InvalidateSnippetMenu();
                    }
                    catch (Exception exception)
                    {
                        Log.Error(
                            "Schnellbausteine wurden geändert, aber die Ribbon-Menüs konnten nicht aktualisiert werden.",
                            exception);
                        MessageBox.Show(
                            owner,
                            "Die Schnellbausteine wurden geändert, aber die Menüs konnten nicht aktualisiert werden.\n\n" +
                                exception.Message,
                            "Menüaktualisierung nicht möglich",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void InvalidateSnippetMenu()
        {
            menu.Invalidate();
            IRibbonUI currentRibbonUi = ribbonUi;
            if (currentRibbonUi == null) return;
            foreach (string controlId in SnippetMenuControlIds)
            {
                Log.Trace("ribbon.invalidate." + controlId, "before");
                currentRibbonUi.InvalidateControl(controlId);
                Log.Trace("ribbon.invalidate." + controlId, "after");
            }
        }

        private void DispatchUserAction(
            string operation,
            Task callbackBodyCompleted,
            string errorMessage,
            Action<IntPtr> action)
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

    internal sealed class SnippetMenuModel
    {
        private static readonly XNamespace Ribbon =
            "http://schemas.microsoft.com/office/2009/07/customui";
        private const string EmptyLabel = "Keine Schnellbausteine gespeichert";
        private const string ManageLabel = "Schnellbausteine verwalten";
        private readonly SnippetRepository repository;
        private readonly object gate = new object();
        private readonly IDictionary<string, IDictionary<string, Snippet>> snapshots =
            new Dictionary<string, IDictionary<string, Snippet>>(StringComparer.Ordinal);
        private bool valid;

        public SnippetMenuModel(SnippetRepository repository)
        {
            if (repository == null) throw new ArgumentNullException("repository");
            this.repository = repository;
        }

        public string RefreshAndGetContent(string menuControlId)
        {
            lock (gate)
            {
                if (string.IsNullOrEmpty(menuControlId))
                {
                    throw new InvalidOperationException(
                        "Die Schnellbaustein-Menüinstanz ist nicht eindeutig.");
                }

                IList<Snippet> snippets = repository.GetAll()
                    .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                string generation = Guid.NewGuid().ToString("N");
                var nextSnapshot = new Dictionary<string, Snippet>(
                    StringComparer.Ordinal);
                var content = new XElement(Ribbon + "menu");
                if (snippets.Count == 0)
                {
                    content.Add(new XElement(
                        Ribbon + "button",
                        new XAttribute("id", "sqEmpty_" + generation),
                        new XAttribute("label", EmptyLabel),
                        new XAttribute("enabled", "false")));
                }
                else
                {
                    for (int index = 0; index < snippets.Count; index++)
                    {
                        string controlId = "sqSnippet_" + generation + "_" + index;
                        nextSnapshot.Add(controlId, snippets[index]);
                        content.Add(new XElement(
                            Ribbon + "button",
                            new XAttribute("id", controlId),
                            new XAttribute("label", snippets[index].Name),
                            new XAttribute("onAction", "InsertSnippet")));
                    }
                }
                content.Add(
                    new XElement(
                        Ribbon + "menuSeparator",
                        new XAttribute("id", "sqManageSeparator_" + generation)),
                    new XElement(
                        Ribbon + "button",
                        new XAttribute("id", "sqManage_" + generation),
                        new XAttribute("label", ManageLabel),
                        new XAttribute("onAction", "ManageSnippets")));
                snapshots[menuControlId] = nextSnapshot;
                valid = true;
                return content.ToString(SaveOptions.DisableFormatting);
            }
        }

        public Snippet GetSnippet(string controlId)
        {
            lock (gate)
            {
                EnsureSnapshot();
                Snippet snippet;
                if (string.IsNullOrEmpty(controlId))
                {
                    throw new InvalidOperationException(
                        "Die ausgewählte Menüposition ist nicht mehr eindeutig.");
                }
                foreach (IDictionary<string, Snippet> snapshot in snapshots.Values)
                {
                    if (snapshot.TryGetValue(controlId, out snippet)) return snippet;
                }
                throw new InvalidOperationException(
                    "Die ausgewählte Menüposition ist nicht mehr eindeutig.");
            }
        }

        public void Invalidate()
        {
            lock (gate)
            {
                valid = false;
                snapshots.Clear();
            }
        }

        private void EnsureSnapshot()
        {
            if (!valid)
            {
                throw new InvalidOperationException(
                    "Das Schnellbaustein-Menü wurde noch nicht aktualisiert.");
            }
        }
    }
}
