using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
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
        private IRibbonUI ribbon;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInstance, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (ribbon != null && Marshal.IsComObject(ribbon)) Marshal.FinalReleaseComObject(ribbon);
            ribbon = null;
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
            try
            {
                XDocument fragment;
                using (var oneNote = new OneNoteClient())
                {
                    fragment = OneNoteXml.ExtractSelection(oneNote.ReadCurrentPage(), oneNote.ReadBinary);
                }

                using (var dialog = new NameDialog("Als Schnellbaustein speichern", string.Empty))
                {
                    while (dialog.ShowDialog() == DialogResult.OK)
                    {
                        if (dialog.SnippetName.Length == 0)
                        {
                            MessageBox.Show("Bitte einen Namen eingeben.", "Name erforderlich", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            continue;
                        }

                        bool overwrite = false;
                        if (repository.ContainsName(dialog.SnippetName, null))
                        {
                            DialogResult answer = MessageBox.Show(
                                "Ein Schnellbaustein mit diesem Namen existiert bereits.\n\nJa: uberschreiben\nNein: anderen Namen eingeben",
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
            catch (Exception exception)
            {
                ShowError("Der Schnellbaustein konnte nicht gespeichert werden.", exception);
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
            try
            {
                if (control == null || string.IsNullOrEmpty(control.Tag)) throw new InvalidOperationException("Der Schnellbaustein besitzt keine ID.");
                XDocument fragment = repository.LoadFragment(control.Tag);
                using (var oneNote = new OneNoteClient())
                {
                    XDocument page = oneNote.ReadCurrentPage();
                    OneNoteXml.Insert(page, fragment);
                    oneNote.UpdateCurrentPage(page);
                }
            }
            catch (Exception exception)
            {
                ShowError("Der Schnellbaustein konnte nicht eingefugt werden.", exception);
            }
        }

        public void ManageSnippets(IRibbonControl control)
        {
            try
            {
                using (var dialog = new ManageDialog(repository, InvalidateMenu))
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception exception)
            {
                ShowError("Die Verwaltung konnte nicht geoffnet werden.", exception);
            }
        }

        private void InvalidateMenu()
        {
            if (ribbon != null) ribbon.InvalidateControl("sqInsertMenu");
        }

        private static void ShowError(string message, Exception exception)
        {
            Log.Error(message, exception);
            MessageBox.Show(message + "\n\n" + exception.Message, "Schnellbausteine", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
