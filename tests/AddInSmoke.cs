using System;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace Schnellbausteine.Tests
{
    internal static class AddInSmoke
    {
        private static int Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.Error.WriteLine("Usage: Schnellbausteine.AddInSmoke.exe <add-in-path>");
                return 1;
            }

            try
            {
                Assembly assembly = Assembly.LoadFrom(args[0]);
                Type type = assembly.GetType("Schnellbausteine.AddIn.AddIn", true);
                object addIn = Activator.CreateInstance(type);
                string ribbonXml = (string)type.GetMethod("GetCustomUI").Invoke(
                    addIn, new object[] { "Microsoft.OneNote" });
                string menuXml = (string)type.GetMethod("GetSnippetMenuContent").Invoke(
                    addIn, new object[] { null });

                XNamespace ribbonNamespace =
                    "http://schemas.microsoft.com/office/2009/07/customui";
                XDocument ribbon = XDocument.Parse(ribbonXml);
                XElement contextMenus = ribbon.Root.Element(ribbonNamespace + "contextMenus");
                XElement textMenu = contextMenus == null
                    ? null
                    : contextMenus.Elements(ribbonNamespace + "contextMenu").SingleOrDefault(
                        element => (string)element.Attribute("idMso") == "ContextMenuText");
                XElement saveButton = textMenu == null
                    ? null
                    : textMenu.Elements(ribbonNamespace + "button").SingleOrDefault(
                        element => (string)element.Attribute("id") == "sqSaveSelection");

                if (!ribbonXml.Contains("sqInsertMenu") ||
                    !ribbonXml.Contains("label=\"Einfügen\"") ||
                    saveButton == null ||
                    (string)saveButton.Attribute("label") != "Als Schnellbaustein speichern..." ||
                    (string)saveButton.Attribute("imageMso") !=
                        "SaveSelectionToQuickPartGallery" ||
                    (string)saveButton.Attribute("onAction") != "SaveSelection" ||
                    (string)saveButton.Attribute("insertBeforeMso") != "Cut" ||
                    saveButton.Attribute("getVisible") != null ||
                    type.GetMethod("SaveSelection") == null ||
                    !menuXml.Contains("sqEmpty"))
                {
                    Console.Error.WriteLine("Ribbon content is incomplete.");
                    return 1;
                }

                Console.WriteLine("Managed add-in and Ribbon smoke test succeeded.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }
    }
}
