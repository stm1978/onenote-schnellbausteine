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
        private IRibbonUI ribbon;
        private int operationActive;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInstance, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            IRibbonUI disconnectedRibbon = Interlocked.Exchange(ref ribbon, null);
            bool dispatcherStopped = dispatcher.Shutdown(TimeSpan.FromSeconds(5));
            if (dispatcherStopped && disconnectedRibbon != null && Marshal.IsComObject(disconnectedRibbon))
            {
                Marshal.FinalReleaseComObject(disconnectedRibbon);
            }
            else if (!dispatcherStopped)
            {
                Log.Info("STA-UI-Thread reagierte beim Beenden nicht; Ribbon-RCW wird nicht vorzeitig freigegeben.");
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Schnellbausteine.AddIn.Ribbon.xml"))
            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            ribbon = ribbonUi;
        }

        public bool CanSaveSelection(IRibbonControl control)
        {
            return true;
        }

        internal bool HasCurrentSelection()
        {
            try
            {
                using (var oneNote = new OneNoteClient())
                {
                    return OneNoteXml.HasSelection(oneNote.ReadCurrentPage());
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
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                DispatchUserAction(
                    callbackBodyCompleted.Task,
                    "Der Schnellbaustein konnte nicht gespeichert werden.",
                    SaveSelectionCore);
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
            }
        }

        private void SaveSelectionCore(IntPtr fallbackOwnerHandle)
        {
            XDocument fragment;
            IntPtr ownerHandle;
            using (var oneNote = new OneNoteClient())
            {
                fragment = OneNoteXml.ExtractSelection(oneNote.ReadCurrentPage(), oneNote.ReadBinary);
                ownerHandle = oneNote.CurrentWindowHandle;
            }

            var owner = new WindowHandleOwner(
                ownerHandle == IntPtr.Zero ? fallbackOwnerHandle : ownerHandle);
            using (var dialog = new NameDialog("Als Schnellbaustein speichern", string.Empty))
            {
                while (dialog.ShowDialog(owner) == DialogResult.OK)
                {
                    if (dialog.SnippetName.Length == 0)
                    {
                        MessageBox.Show(owner, "Bitte einen Namen eingeben.", "Name erforderlich",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    bool overwrite = false;
                    if (repository.ContainsName(dialog.SnippetName, null))
                    {
                        DialogResult answer = MessageBox.Show(owner,
                            "Ein Schnellbaustein mit diesem Namen existiert bereits.\n\nJa: überschreiben\nNein: anderen Namen eingeben",
                            "Namenskonflikt", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                        if (answer == DialogResult.Cancel) return;
                        if (answer == DialogResult.No) continue;
                        overwrite = true;
                    }

                    repository.Save(dialog.SnippetName, fragment, overwrite);
                    InvalidateMenu();
                    return;
                }
            }
        }

        public string GetSnippetMenuContent(IRibbonControl control)
        {
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
                return menu.ToString(SaveOptions.DisableFormatting);
            }
            catch (Exception exception)
            {
                Log.Error("Dynamisches Menu", exception);
                return "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><button id=\"sqError\" label=\"Fehler beim Laden\" enabled=\"false\"/></menu>";
            }
        }

        public void InsertSnippet(IRibbonControl control)
        {
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
                    callbackBodyCompleted.Task,
                    "Der Schnellbaustein konnte nicht eingefügt werden.",
                    delegate
                    {
                        if (string.IsNullOrEmpty(snippetId))
                        {
                            throw new InvalidOperationException("Der Schnellbaustein besitzt keine ID.");
                        }
                        XDocument fragment = repository.LoadFragment(snippetId);
                        using (var oneNote = new OneNoteClient())
                        {
                            XDocument page = oneNote.ReadCurrentPage();
                            OneNoteXml.Insert(page, fragment);
                            oneNote.UpdateCurrentPage(page);
                        }
                    });
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
            }
        }

        public void ManageSnippets(IRibbonControl control)
        {
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            try
            {
                DispatchUserAction(
                    callbackBodyCompleted.Task,
                    "Die Verwaltung konnte nicht geöffnet werden.",
                    delegate(IntPtr ownerHandle)
                    {
                        using (var dialog = new ManageDialog(repository, InvalidateMenu))
                        {
                            dialog.ShowDialog(new WindowHandleOwner(ownerHandle));
                        }
                    });
            }
            finally
            {
                callbackBodyCompleted.SetResult(null);
            }
        }

        private void DispatchUserAction(
            Task callbackBodyCompleted, string errorMessage, Action<IntPtr> action)
        {
            if (Interlocked.CompareExchange(ref operationActive, 1, 0) != 0)
            {
                Log.Info("Ein Schnellbausteine-Befehl wird bereits ausgeführt.");
                return;
            }

            IntPtr ownerHandle = NativeMethods.GetForegroundWindow();
            try
            {
                dispatcher.PostAfterCallback(
                    callbackBodyCompleted,
                    delegate
                    {
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
        }

        private void InvalidateMenu()
        {
            IRibbonUI activeRibbon = ribbon;
            if (activeRibbon != null) activeRibbon.InvalidateControl("sqInsertMenu");
        }

        private static void ShowError(string message, Exception exception, IntPtr ownerHandle)
        {
            Log.Error(message, exception);
            MessageBox.Show(new WindowHandleOwner(ownerHandle), message + "\n\n" + exception.Message,
                "Schnellbausteine", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll")]
            internal static extern IntPtr GetForegroundWindow();
        }
    }
}
