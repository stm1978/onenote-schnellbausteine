using System;
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
        private IRibbonUI ribbon;
        private int operationActive;

        public AddIn()
        {
            saveSelectionAction = SaveSelectionCore;
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
            IRibbonUI disconnectedRibbon = Interlocked.Exchange(ref ribbon, null);
            bool dispatcherStopped = dispatcher.Shutdown(TimeSpan.FromSeconds(5));
            if (dispatcherStopped && disconnectedRibbon != null && Marshal.IsComObject(disconnectedRibbon))
            {
                Log.Trace("ribbon.release", "before");
                Marshal.FinalReleaseComObject(disconnectedRibbon);
                Log.Trace("ribbon.release", "after");
            }
            else if (!dispatcherStopped)
            {
                Log.Info("STA-UI-Thread reagierte beim Beenden nicht; Ribbon-RCW wird nicht vorzeitig freigegeben.");
            }
            Log.Trace("addin.disconnection", "exit");
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
            ribbon = ribbonUi;
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
                    InvalidateMenu();
                    Log.Trace("command.SaveSelection", "exit-saved");
                    return;
                }
            }
            Log.Trace("command.SaveSelection", "exit-cancelled");
        }

        public string GetSnippetMenuContent(IRibbonControl control)
        {
            Log.Trace("callback.GetSnippetMenuContent", "entry");
            try
            {
                const string ns = "http://schemas.microsoft.com/office/2009/07/customui";
                var menu = new XElement(XName.Get("menu", ns));
                foreach (Snippet item in repository.GetAll())
                {
                    menu.Add(new XElement(XName.Get("button", ns),
                        new XAttribute("id", "sqSnippet_" + item.Id),
                        new XAttribute("label", item.Name),
                        new XAttribute("tag", item.Id),
                        new XAttribute("onAction", "InsertSnippet")));
                }
                if (!menu.HasElements)
                {
                    menu.Add(new XElement(XName.Get("button", ns),
                        new XAttribute("id", "sqEmpty"),
                        new XAttribute("label", "Keine Schnellbausteine gespeichert"),
                        new XAttribute("enabled", "false")));
                }
                string xml = menu.ToString(SaveOptions.DisableFormatting);
                Log.Trace("callback.GetSnippetMenuContent", "exit");
                return xml;
            }
            catch (Exception exception)
            {
                Log.Error("Dynamisches Menu", exception);
                return "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><button id=\"sqError\" label=\"Fehler beim Laden\" enabled=\"false\"/></menu>";
            }
        }

        public void InsertSnippet(IRibbonControl control)
        {
            Log.Trace("callback.InsertSnippet", "entry");
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                string snippetId;
                try
                {
                    snippetId = control == null ? null : control.Tag;
                }
                catch (Exception exception)
                {
                    DispatchUserAction(
                        "InsertSnippet",
                        callbackBodyCompleted.Task,
                        "Der Schnellbaustein konnte nicht eingefügt werden.",
                        delegate
                        {
                            throw new InvalidOperationException(
                                "Die ID des Schnellbausteins konnte nicht gelesen werden.", exception);
                        });
                    return;
                }
                DispatchUserAction(
                    "InsertSnippet",
                    callbackBodyCompleted.Task,
                    "Der Schnellbaustein konnte nicht eingefügt werden.",
                    delegate
                    {
                        Log.Trace("command.InsertSnippet", "entry");
                        if (string.IsNullOrEmpty(snippetId))
                        {
                            throw new InvalidOperationException("Der Schnellbaustein besitzt keine ID.");
                        }
                        Log.Trace("repository.load", "before");
                        XDocument fragment = repository.LoadFragment(snippetId);
                        Log.Trace("repository.load", "after");
                        using (var oneNote = new OneNoteClient())
                        {
                            XDocument page = oneNote.ReadCurrentPage();
                            Log.Trace("onenote.xml.insert", "before");
                            OneNoteXml.Insert(page, fragment);
                            Log.Trace("onenote.xml.insert", "after");
                            oneNote.UpdateCurrentPage(page);
                        }
                        Log.Trace("command.InsertSnippet", "exit");
                    });
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
                Log.Trace("callback.InsertSnippet", "exit");
            }
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
                    "Die Verwaltung konnte nicht geöffnet werden.",
                    delegate(IntPtr ownerHandle)
                    {
                        using (var dialog = new ManageDialog(repository, InvalidateMenu))
                        {
                            Log.Trace("ui.ManageDialog", "before");
                            dialog.ShowDialog(new WindowHandleOwner(ownerHandle));
                            Log.Trace("ui.ManageDialog", "after");
                        }
                    });
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
                Log.Trace("callback.ManageSnippets", "exit");
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

        private void InvalidateMenu()
        {
            IRibbonUI activeRibbon = ribbon;
            if (activeRibbon != null)
            {
                Log.Trace("ribbon.InvalidateControl", "before");
                activeRibbon.InvalidateControl("sqInsertMenu");
                Log.Trace("ribbon.InvalidateControl", "after");
            }
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
