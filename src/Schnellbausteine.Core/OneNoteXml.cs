using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Schnellbausteine.Core
{
    public static class OneNoteXml
    {
        public static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        public static readonly XNamespace Snippet = "urn:schnellbausteine:v1";

        public static bool HasSelection(XDocument page)
        {
            if (page == null) throw new ArgumentNullException("page");
            List<XElement> selected = page.Descendants().Where(IsSelected).ToList();
            if (selected.Count == 0) return false;
            if (selected.Any(element => element.Name != One + "T")) return true;
            if (selected.Count > 1) return selected.Any(element => element.Value.Length > 0);
            return selected[0].Value.Length > 0 && !ContainsSpecialMarkup(selected[0]);
        }

        public static XDocument ExtractSelection(XDocument page, Func<string, string> binaryResolver)
        {
            if (page == null) throw new ArgumentNullException("page");
            if (!HasSelection(page)) throw new InvalidOperationException("OneNote meldet keine markierte Auswahl.");
            var selected = page.Descendants().Where(IsSelected).ToList();

            var roots = new List<XElement>();
            foreach (XElement marker in selected)
            {
                XElement candidate = marker.Name == One + "OE" ? marker : marker.Ancestors(One + "OE").FirstOrDefault();
                if (candidate == null) continue;

                XElement richAncestor = marker.Ancestors().FirstOrDefault(element =>
                    element.Name == One + "Table" || element.Name == One + "Image");
                if (richAncestor != null)
                {
                    candidate = richAncestor.Ancestors(One + "OE").FirstOrDefault() ?? candidate;
                }

                if (!roots.Any(root => ReferenceEquals(root, candidate)) &&
                    !roots.Any(root => candidate.Ancestors().Any(ancestor => ReferenceEquals(ancestor, root))))
                {
                    roots.RemoveAll(root => root.Ancestors().Any(ancestor => ReferenceEquals(ancestor, candidate)));
                    roots.Add(candidate);
                }
            }

            if (roots.Count == 0) throw new InvalidOperationException("Die OneNote-Auswahl enthalt keinen speicherbaren Inhalt.");

            var content = new XElement(Snippet + "Content");
            foreach (XElement root in roots)
            {
                XElement clone = new XElement(root);
                PruneUnselectedText(clone);
                RemoveTransientAttributes(clone);
                ResolveImageData(clone, binaryResolver);
                content.Add(clone);
            }

            var usedTagIndexes = new HashSet<string>(
                content.Descendants(One + "Tag")
                    .Select(tag => (string)tag.Attribute("index"))
                    .Where(index => !string.IsNullOrEmpty(index)),
                StringComparer.Ordinal);

            var definitions = new XElement(Snippet + "TagDefs");
            var quickStyles = new XElement(Snippet + "QuickStyleDefs");
            XElement pageRoot = page.Root;
            if (pageRoot != null)
            {
                foreach (XElement definition in pageRoot.Elements(One + "TagDef"))
                {
                    if (usedTagIndexes.Contains((string)definition.Attribute("index")))
                    {
                        XElement clone = new XElement(definition);
                        RemoveTransientAttributes(clone);
                        definitions.Add(clone);
                    }
                }

                var usedQuickStyles = new HashSet<string>(
                    content.DescendantsAndSelf()
                        .Select(element => (string)element.Attribute("quickStyleIndex"))
                        .Where(index => !string.IsNullOrEmpty(index)),
                    StringComparer.Ordinal);
                foreach (XElement definition in pageRoot.Elements(One + "QuickStyleDef"))
                {
                    if (usedQuickStyles.Contains((string)definition.Attribute("index")))
                    {
                        XElement clone = new XElement(definition);
                        RemoveTransientAttributes(clone);
                        quickStyles.Add(clone);
                    }
                }
            }

            return new XDocument(
                new XElement(Snippet + "Snippet",
                    new XAttribute("version", "1"),
                    definitions,
                    quickStyles,
                    content));
        }

        public static void Insert(XDocument page, XDocument fragment)
        {
            if (page == null) throw new ArgumentNullException("page");
            if (fragment == null) throw new ArgumentNullException("fragment");
            ValidateFragment(fragment);

            List<XElement> incoming = fragment.Root.Element(Snippet + "Content").Elements().Select(element => new XElement(element)).ToList();
            if (incoming.Count == 0) throw new InvalidOperationException("Der Schnellbaustein ist leer.");

            foreach (XElement element in incoming) RemoveTransientAttributes(element);
            MergeTagDefinitions(page, fragment, incoming);
            MergeQuickStyleDefinitions(page, fragment, incoming);

            List<XElement> markers = page.Descendants().Where(IsSelected).ToList();
            if (markers.Count == 0)
            {
                InsertAtOutlineEnd(page, incoming);
                return;
            }

            bool specialCursor = markers.Count == 1 && markers[0].Name == One + "T" &&
                ContainsSpecialMarkup(markers[0]);
            if (specialCursor)
            {
                InsertAfterContainingParagraph(markers[0], incoming);
                return;
            }

            if (CanInsertInline(incoming) &&
                markers.All(marker => marker.Name == One + "T") &&
                markers.Select(marker => marker.Parent).Distinct().Count() == 1)
            {
                InsertInline(markers, incoming[0]);
                return;
            }

            InsertBlocks(markers, incoming);
        }

        private static void InsertInline(IList<XElement> markers, XElement incoming)
        {
            XElement first = markers[0];
            List<XNode> nodes = incoming.Nodes().Select(CloneNode).ToList();
            foreach (XElement marker in markers)
            {
                XAttribute selected = marker.Attribute("selected");
                if (selected != null) selected.Remove();
                if (ReferenceEquals(marker, first))
                {
                    marker.ReplaceWith(nodes);
                }
                else
                {
                    marker.Remove();
                }
            }

            XElement lastText = nodes.OfType<XElement>().LastOrDefault(element => element.Name == One + "T");
            if (lastText != null)
            {
                lastText.AddAfterSelf(new XElement(One + "T",
                    new XAttribute("selected", "all"),
                    new XCData(string.Empty)));
            }
        }

        private static void InsertBlocks(IList<XElement> markers, IList<XElement> incoming)
        {
            XElement first = markers[0];
            XElement paragraph = first.Name == One + "OE" ? first : first.Ancestors(One + "OE").FirstOrDefault();
            if (paragraph == null || paragraph.Parent == null)
            {
                InsertAtOutlineEnd(first.Document, incoming);
                return;
            }

            XElement after = null;
            if (first.Name == One + "T" && ReferenceEquals(first.Parent, paragraph))
            {
                after = SplitAfterSelection(paragraph, markers);
            }

            List<XElement> affectedParagraphs = markers
                .Select(marker => marker.Name == One + "OE" ? marker : marker.Ancestors(One + "OE").FirstOrDefault())
                .Where(item => item != null)
                .Distinct()
                .ToList();
            foreach (XElement marker in markers)
            {
                XAttribute selected = marker.Attribute("selected");
                if (selected != null) selected.Remove();
                if (marker.Name == One + "T")
                {
                    if (marker.Parent != null) marker.Remove();
                }
                else if (marker.Name == One + "OE" && marker.Parent != null && !ReferenceEquals(marker, paragraph))
                {
                    marker.Remove();
                }
                else if (marker.Name != One + "OE" && marker.Parent != null)
                {
                    marker.Remove();
                }
            }

            object[] insertion = incoming.Cast<object>().ToArray();
            paragraph.AddAfterSelf(insertion);
            XElement last = incoming.Last();
            if (after != null && after.Nodes().Any())
            {
                last.AddAfterSelf(after);
                SetCaretAtStart(after);
            }
            else
            {
                SetCaretAfter(last);
            }

            if (first.Name == One + "OE" || !paragraph.Nodes().Any()) paragraph.Remove();
            foreach (XElement affected in affectedParagraphs.Where(item =>
                !ReferenceEquals(item, paragraph) && item.Parent != null && !item.Nodes().Any()).ToList())
            {
                affected.Remove();
            }
        }

        private static XElement SplitAfterSelection(XElement paragraph, IList<XElement> markers)
        {
            var direct = paragraph.Elements().ToList();
            int lastSelected = direct.FindLastIndex(element => markers.Contains(element));
            if (lastSelected < 0 || lastSelected == direct.Count - 1) return null;

            var after = new XElement(One + "OE");
            CopyFormattingAttributes(paragraph, after);
            foreach (XElement element in direct.Skip(lastSelected + 1).ToList())
            {
                element.Remove();
                after.Add(element);
            }
            return after;
        }

        private static void InsertAfterContainingParagraph(XElement marker, IList<XElement> incoming)
        {
            XAttribute selected = marker.Attribute("selected");
            if (selected != null) selected.Remove();
            XElement paragraph = marker.Ancestors(One + "OE").FirstOrDefault();
            if (paragraph == null || paragraph.Parent == null)
            {
                InsertAtOutlineEnd(marker.Document, incoming);
                return;
            }

            paragraph.AddAfterSelf(incoming.Cast<object>().ToArray());
            SetCaretAfter(incoming.Last());
        }

        private static void InsertAtOutlineEnd(XDocument page, IList<XElement> incoming)
        {
            XElement pageRoot = page.Root;
            if (pageRoot == null) throw new InvalidOperationException("Das OneNote-Seitendokument besitzt kein Wurzelelement.");
            XElement children = page.Descendants(One + "OEChildren").FirstOrDefault();
            if (children == null)
            {
                var outline = new XElement(One + "Outline", new XElement(One + "OEChildren"));
                pageRoot.Add(outline);
                children = outline.Element(One + "OEChildren");
            }
            children.Add(incoming);
            SetCaretAfter(incoming.Last());
        }

        private static void SetCaretAtStart(XElement paragraph)
        {
            ClearSelection(paragraph.Document);
            paragraph.AddFirst(new XElement(One + "T",
                new XAttribute("selected", "all"),
                new XCData(string.Empty)));
        }

        private static void SetCaretAfter(XElement element)
        {
            ClearSelection(element.Document);
            var caret = new XElement(One + "OE",
                new XElement(One + "T",
                    new XAttribute("selected", "all"),
                    new XCData(string.Empty)));
            element.AddAfterSelf(caret);
        }

        private static void ClearSelection(XDocument document)
        {
            foreach (XAttribute selected in document.Descendants().Attributes("selected").ToList()) selected.Remove();
        }

        private static void MergeTagDefinitions(XDocument page, XDocument fragment, IList<XElement> content)
        {
            XElement pageRoot = page.Root;
            if (pageRoot == null) throw new InvalidOperationException("Das OneNote-Seitendokument besitzt kein Wurzelelement.");

            List<XElement> targetDefinitions = pageRoot.Elements(One + "TagDef").ToList();
            int next = targetDefinitions
                .Select(definition => ParseIndex((string)definition.Attribute("index")))
                .DefaultIfEmpty(-1).Max() + 1;
            var map = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (XElement source in fragment.Root.Element(Snippet + "TagDefs").Elements(One + "TagDef"))
            {
                string oldIndex = (string)source.Attribute("index");
                if (string.IsNullOrEmpty(oldIndex)) continue;
                XElement match = targetDefinitions.FirstOrDefault(target => TagSignature(target) == TagSignature(source));
                if (match == null)
                {
                    match = new XElement(source);
                    match.SetAttributeValue("index", next.ToString());
                    pageRoot.AddFirst(match);
                    targetDefinitions.Add(match);
                    next++;
                }
                map[oldIndex] = (string)match.Attribute("index");
            }

            foreach (XElement tag in content.SelectMany(item => item.Descendants(One + "Tag")))
            {
                string index = (string)tag.Attribute("index");
                string mapped;
                if (!string.IsNullOrEmpty(index) && map.TryGetValue(index, out mapped))
                {
                    tag.SetAttributeValue("index", mapped);
                }
            }
        }

        private static void MergeQuickStyleDefinitions(XDocument page, XDocument fragment, IList<XElement> content)
        {
            XElement sourceContainer = fragment.Root.Element(Snippet + "QuickStyleDefs");
            if (sourceContainer == null) return;

            XElement pageRoot = page.Root;
            List<XElement> targetDefinitions = pageRoot.Elements(One + "QuickStyleDef").ToList();
            int next = targetDefinitions
                .Select(definition => ParseIndex((string)definition.Attribute("index")))
                .DefaultIfEmpty(-1).Max() + 1;
            var map = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (XElement source in sourceContainer.Elements(One + "QuickStyleDef"))
            {
                string oldIndex = (string)source.Attribute("index");
                if (string.IsNullOrEmpty(oldIndex)) continue;
                XElement match = targetDefinitions.FirstOrDefault(target => DefinitionSignature(target) == DefinitionSignature(source));
                if (match == null)
                {
                    match = new XElement(source);
                    match.SetAttributeValue("index", next.ToString());
                    XElement lastDefinition = pageRoot.Elements()
                        .Where(element => element.Name == One + "TagDef" || element.Name == One + "QuickStyleDef")
                        .LastOrDefault();
                    if (lastDefinition == null) pageRoot.AddFirst(match);
                    else lastDefinition.AddAfterSelf(match);
                    targetDefinitions.Add(match);
                    next++;
                }
                map[oldIndex] = (string)match.Attribute("index");
            }

            foreach (XElement element in content.SelectMany(item => item.DescendantsAndSelf()))
            {
                string index = (string)element.Attribute("quickStyleIndex");
                string mapped;
                if (!string.IsNullOrEmpty(index) && map.TryGetValue(index, out mapped))
                {
                    element.SetAttributeValue("quickStyleIndex", mapped);
                }
            }
        }

        private static string TagSignature(XElement definition)
        {
            return DefinitionSignature(definition);
        }

        private static string DefinitionSignature(XElement definition)
        {
            return string.Join("|", definition.Attributes()
                .Where(attribute => attribute.Name.LocalName != "index" && attribute.Name.LocalName != "objectID")
                .OrderBy(attribute => attribute.Name.ToString())
                .Select(attribute => attribute.Name + "=" + attribute.Value));
        }

        private static int ParseIndex(string value)
        {
            int result;
            return int.TryParse(value, out result) ? result : -1;
        }

        private static bool CanInsertInline(IList<XElement> incoming)
        {
            if (incoming.Count != 1 || incoming[0].Name != One + "OE") return false;
            return !incoming[0].HasAttributes &&
                incoming[0].Elements().All(element => element.Name == One + "T");
        }

        private static bool ContainsSpecialMarkup(XElement text)
        {
            string value = text.Value;
            return value.IndexOf("<a ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("<!--", StringComparison.Ordinal) >= 0 ||
                   text.Descendants().Any();
        }

        private static void ResolveImageData(XElement content, Func<string, string> binaryResolver)
        {
            foreach (XElement image in content.DescendantsAndSelf(One + "Image"))
            {
                if (image.Element(One + "Data") != null) continue;
                string callback = (string)image.Attribute("callbackID");
                if (string.IsNullOrWhiteSpace(callback))
                {
                    throw new InvalidOperationException("Ein Bild enthalt weder Data noch callbackID.");
                }
                if (binaryResolver == null) throw new InvalidOperationException("Fur ein Callback-Bild ist kein Binardatenleser verfugbar.");
                string data = binaryResolver(callback);
                if (string.IsNullOrWhiteSpace(data)) throw new InvalidOperationException("OneNote lieferte keine Binardaten fur Bild " + callback + ".");
                image.Add(new XElement(One + "Data", data));
                XAttribute callbackAttribute = image.Attribute("callbackID");
                if (callbackAttribute != null) callbackAttribute.Remove();
            }
        }

        private static void PruneUnselectedText(XElement clone)
        {
            List<XElement> selectedTexts = clone.DescendantsAndSelf(One + "T").Where(IsSelected).ToList();
            if (selectedTexts.Count == 0) return;
            foreach (XElement text in clone.DescendantsAndSelf(One + "T").Where(text => !IsSelected(text)).ToList())
            {
                text.Remove();
            }
        }

        private static void RemoveTransientAttributes(XElement element)
        {
            foreach (XAttribute attribute in element.DescendantsAndSelf().Attributes().Where(attribute =>
                attribute.Name.LocalName == "objectID" ||
                attribute.Name.LocalName == "selected" ||
                attribute.Name.LocalName == "lastModifiedTime").ToList())
            {
                attribute.Remove();
            }
        }

        private static bool IsSelected(XElement element)
        {
            string value = (string)element.Attribute("selected");
            return string.Equals(value, "all", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "partial", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static void ValidateFragment(XDocument fragment)
        {
            if (fragment.Root == null || fragment.Root.Name != Snippet + "Snippet" ||
                (string)fragment.Root.Attribute("version") != "1" ||
                fragment.Root.Element(Snippet + "TagDefs") == null ||
                fragment.Root.Element(Snippet + "Content") == null)
            {
                throw new InvalidOperationException("Unbekanntes Schnellbaustein-Format.");
            }
        }

        private static XNode CloneNode(XNode node)
        {
            var element = node as XElement;
            if (element != null) return new XElement(element);
            var cdata = node as XCData;
            if (cdata != null) return new XCData(cdata.Value);
            var text = node as XText;
            if (text != null) return new XText(text.Value);
            var comment = node as XComment;
            if (comment != null) return new XComment(comment.Value);
            var processingInstruction = node as XProcessingInstruction;
            if (processingInstruction != null)
                return new XProcessingInstruction(processingInstruction.Target, processingInstruction.Data);
            throw new InvalidOperationException("Nicht unterstutzter XML-Knotentyp: " + node.NodeType);
        }

        private static void CopyFormattingAttributes(XElement source, XElement target)
        {
            foreach (XAttribute attribute in source.Attributes().Where(attribute =>
                attribute.Name.LocalName == "style" ||
                attribute.Name.LocalName == "quickStyleIndex" ||
                attribute.Name.LocalName == "lang"))
            {
                target.SetAttributeValue(attribute.Name, attribute.Value);
            }
        }
    }
}
