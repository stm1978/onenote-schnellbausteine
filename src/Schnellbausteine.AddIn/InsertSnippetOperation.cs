using System;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    internal static class InsertSnippetOperation
    {
        public static void Execute(
            OneNoteClient oneNote,
            XDocument anchoredPage,
            Snippet snippet,
            Func<string, XDocument> loadFragment)
        {
            if (oneNote == null) throw new ArgumentNullException("oneNote");
            if (anchoredPage == null) throw new ArgumentNullException("anchoredPage");
            if (snippet == null) throw new ArgumentNullException("snippet");
            if (loadFragment == null) throw new ArgumentNullException("loadFragment");

            XElement target = GetTargetOutline(anchoredPage);
            OneNoteClient.ValidateInkFreeUpdate(new XDocument(new XElement(target)));
            Stopwatch stopwatch = Stopwatch.StartNew();
            Log.Trace("repository.load", "before");
            XDocument fragment = loadFragment(snippet.Id);
            Log.Trace("repository.load", "after");
            OneNoteClient.ValidateInkFreeUpdate(fragment);
            Log.Diagnostic("performance.insert.fragment-load", "elapsedMs=" + stopwatch.ElapsedMilliseconds);
            stopwatch.Restart();
            Log.Trace("onenote.xml.insert", "before");
            XDocument workingPage = new XDocument(anchoredPage);
            OneNoteInsertionResult insertion =
                OneNoteXml.InsertAtSelection(workingPage, fragment);
            XDocument update = CreateScopedUpdate(anchoredPage, workingPage, target);
            Log.Diagnostic("onenote.insert.anchor", insertion.ToDiagnosticString());
            Log.Trace("onenote.xml.insert", "after");
            Log.Diagnostic("performance.insert.xml-transform", "elapsedMs=" + stopwatch.ElapsedMilliseconds);
            oneNote.UpdateCurrentPageAndActivateWindow(update);
        }

        private static XElement GetTargetOutline(XDocument page)
        {
            if (page.Root == null || page.Root.Name != OneNoteXml.One + "Page" ||
                string.IsNullOrWhiteSpace((string)page.Root.Attribute("ID")))
            {
                throw new InvalidOperationException("Die Zielseite besitzt keine eindeutige Seiten-ID.");
            }
            XElement[] targets = page.Descendants()
                .Where(element => element.Attributes("selected").Any(attribute =>
                    string.Equals(attribute.Value, "all", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Value, "partial", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Value, "true", StringComparison.OrdinalIgnoreCase)))
                .Select(element => element.AncestorsAndSelf()
                    .FirstOrDefault(ancestor => ancestor.Parent == page.Root))
                .Distinct().ToArray();
            if (targets.Length != 1 || targets[0] == null ||
                targets[0].Name != OneNoteXml.One + "Outline" ||
                string.IsNullOrWhiteSpace((string)targets[0].Attribute("objectID")) ||
                page.Root.Elements().Count(element =>
                    (string)element.Attribute("objectID") ==
                    (string)targets[0].Attribute("objectID")) != 1)
            {
                throw new InvalidOperationException(
                    "Die Auswahl besitzt keinen eindeutig identifizierten Textcontainer. Die Seite wurde nicht aktualisiert.");
            }
            return targets[0];
        }

        internal static XDocument CreateScopedUpdate(
            XDocument original, XDocument modified, XElement originalTarget)
        {
            string targetId = (string)originalTarget.Attribute("objectID");
            XElement[] targets = modified.Root.Elements(OneNoteXml.One + "Outline")
                .Where(element => (string)element.Attribute("objectID") == targetId).ToArray();
            if (targets.Length != 1)
                throw new InvalidOperationException("Der Zielcontainer wurde beim Einfügen unerwartet verändert.");

            XElement[] before = original.Root.Elements()
                .Where(element => element != originalTarget && !IsDefinition(element)).ToArray();
            XElement[] after = modified.Root.Elements()
                .Where(element => element != targets[0] && !IsDefinition(element)).ToArray();
            if (before.Length != after.Length ||
                before.Where((element, index) => !EqualIgnoringSelection(element, after[index])).Any())
            {
                throw new InvalidOperationException(
                    "Einfügen würde weitere Seitenobjekte verändern. Die Seite wurde nicht aktualisiert.");
            }

            var update = new XElement(original.Root.Name,
                original.Root.Attributes().Where(attribute => attribute.IsNamespaceDeclaration),
                new XAttribute("ID", (string)original.Root.Attribute("ID")));
            foreach (XElement definition in modified.Root.Elements().Where(IsDefinition))
            {
                XElement[] existing = original.Root.Elements(definition.Name)
                    .Where(element => (string)element.Attribute("index") ==
                        (string)definition.Attribute("index")).ToArray();
                if (existing.Length > 1 || existing.Length == 1 &&
                    !XNode.DeepEquals(existing[0], definition))
                {
                    throw new InvalidOperationException(
                        "Bestehende Formatdefinitionen dürfen beim Einfügen nicht überschrieben werden.");
                }
                if (existing.Length == 0) update.Add(new XElement(definition));
            }
            if (original.Root.Elements().Where(IsDefinition).Any(definition =>
                !modified.Root.Elements(definition.Name).Any(element =>
                    XNode.DeepEquals(element, definition))))
            {
                throw new InvalidOperationException("Eine bestehende Formatdefinition fehlt nach dem Einfügen.");
            }
            update.Add(new XElement(targets[0]));
            var result = new XDocument(update);
            OneNoteClient.ValidateInkFreeUpdate(result);
            OneNoteXml.ValidatePostInsertCaret(result);
            return result;
        }

        private static bool IsDefinition(XElement element)
        {
            return element.Name == OneNoteXml.One + "TagDef" ||
                element.Name == OneNoteXml.One + "QuickStyleDef";
        }

        private static bool EqualIgnoringSelection(XElement first, XElement second)
        {
            var left = new XElement(first);
            var right = new XElement(second);
            left.DescendantsAndSelf().Attributes("selected").Remove();
            right.DescendantsAndSelf().Attributes("selected").Remove();
            return XNode.DeepEquals(left, right);
        }
    }
}
