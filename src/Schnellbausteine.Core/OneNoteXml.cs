using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Schnellbausteine.Core
{
    public sealed class OneNoteSelectionInfo
    {
        internal OneNoteSelectionInfo(
            string scope,
            string containerScope,
            int markerCount,
            string markerTypes,
            string markerPaths,
            string anchorPath,
            bool isUsable,
            string problem)
        {
            Scope = scope;
            ContainerScope = containerScope;
            MarkerCount = markerCount;
            MarkerTypes = markerTypes;
            MarkerPaths = markerPaths;
            AnchorPath = anchorPath;
            IsUsable = isUsable;
            Problem = problem;
        }

        public string Scope { get; private set; }
        public string ContainerScope { get; private set; }
        public int MarkerCount { get; private set; }
        public string MarkerTypes { get; private set; }
        public string MarkerPaths { get; private set; }
        public string AnchorPath { get; private set; }
        public bool IsUsable { get; private set; }
        public string Problem { get; private set; }

        public string ToDiagnosticString()
        {
            return "scope=" + Scope +
                " container=" + ContainerScope +
                " markers=" + MarkerCount +
                " types=" + MarkerTypes +
                " paths=" + MarkerPaths +
                " anchor=" + AnchorPath +
                " usable=" + IsUsable;
        }
    }

    public sealed class OneNoteInsertionResult
    {
        internal OneNoteInsertionResult(OneNoteSelectionInfo selection, string strategy)
        {
            Selection = selection;
            Strategy = strategy;
        }

        public OneNoteSelectionInfo Selection { get; private set; }
        public string Strategy { get; private set; }

        public string ToDiagnosticString()
        {
            return Selection.ToDiagnosticString() + " strategy=" + Strategy;
        }
    }

    public static class OneNoteXml
    {
        public static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        public static readonly XNamespace Snippet = "urn:schnellbausteine:v1";

        private sealed class SelectionContext
        {
            public List<XElement> Markers;
            public OneNoteSelectionInfo Info;
        }

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
            SelectionContext selection = CreateSelectionContext(page);
            if (!selection.Info.IsUsable)
            {
                throw new InvalidOperationException(selection.Info.Problem);
            }
            if (!HasSelection(page)) throw new InvalidOperationException("OneNote meldet keine markierte Auswahl.");
            List<XElement> selected = selection.Markers;

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

            if (roots.Count == 0) throw new InvalidOperationException("Die OneNote-Auswahl enthält keinen speicherbaren Inhalt.");

            var content = new XElement(Snippet + "Content");
            foreach (XElement root in roots)
            {
                XElement clone = new XElement(root);
                PruneUnselectedContent(clone, false);
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

        public static OneNoteSelectionInfo AnalyzeSelection(XDocument page)
        {
            if (page == null) throw new ArgumentNullException("page");
            return CreateSelectionContext(page).Info;
        }

        public static void Insert(XDocument page, XDocument fragment)
        {
            InsertAtSelection(page, fragment);
        }

        public static OneNoteInsertionResult InsertAtSelection(XDocument page, XDocument fragment)
        {
            if (page == null) throw new ArgumentNullException("page");
            if (fragment == null) throw new ArgumentNullException("fragment");
            ValidateFragment(fragment);

            SelectionContext selection = CreateSelectionContext(page);
            if (!selection.Info.IsUsable)
            {
                throw new InvalidOperationException(selection.Info.Problem);
            }

            List<XElement> incoming = fragment.Root.Element(Snippet + "Content").Elements().Select(element => new XElement(element)).ToList();
            if (incoming.Count == 0) throw new InvalidOperationException("Der Schnellbaustein ist leer.");

            foreach (XElement element in incoming) RemoveTransientAttributes(element);
            MergeTagDefinitions(page, fragment, incoming);
            MergeQuickStyleDefinitions(page, fragment, incoming);
            ClearSelection(page);

            List<XElement> markers = selection.Markers;
            if (selection.Info.Scope == "TableCell")
            {
                InsertIntoTableCell(markers[0], incoming);
                return CompleteInsertion(page, selection.Info, "ReplaceTableCell");
            }

            if (selection.Info.Scope == "SpecialCursor")
            {
                InsertAfterContainingParagraph(markers[0], incoming);
                return CompleteInsertion(page, selection.Info, "AfterSpecialParagraph");
            }

            if (CanInsertInline(incoming) &&
                markers.All(marker => marker.Name == One + "T") &&
                markers.Select(marker => marker.Parent).Distinct().Count() == 1)
            {
                InsertInline(markers, incoming[0]);
                return CompleteInsertion(page, selection.Info, "Inline");
            }

            InsertBlocks(markers, incoming);
            return CompleteInsertion(page, selection.Info, "BlockRange");
        }

        public static void ValidatePostInsertCaret(XDocument page)
        {
            if (page == null) throw new ArgumentNullException("page");
            if (page.Root == null)
            {
                throw new InvalidOperationException("Das OneNote-Seitendokument besitzt kein Wurzelelement.");
            }

            List<XAttribute> selectionMarkers = page.Root
                .DescendantsAndSelf()
                .Attributes()
                .Where(attribute => attribute.Name.LocalName == "selected")
                .ToList();
            if (selectionMarkers.Count != 1)
            {
                throw new InvalidOperationException(
                    "Die aktualisierte Seite muss genau einen Cursoranker enthalten.");
            }

            XAttribute marker = selectionMarkers[0];
            XElement caret = marker.Parent;
            bool validCaret =
                caret != null &&
                caret.Name == One + "T" &&
                string.Equals(marker.Value, "all", StringComparison.Ordinal) &&
                caret.Parent != null &&
                caret.Parent.Name == One + "OE" &&
                caret.Nodes().Count() == 1 &&
                caret.Nodes().Single() is XCData &&
                caret.Value.Length == 0;
            if (!validCaret)
            {
                throw new InvalidOperationException(
                    "Der Cursoranker muss ein leeres one:T selected=\"all\" in einem one:OE sein.");
            }
        }

        public static string GetPostInsertFocusTargetObjectId(XDocument page)
        {
            if (page == null) throw new ArgumentNullException("page");
            if (page.Root == null)
            {
                throw new InvalidOperationException("Das OneNote-Seitendokument besitzt kein Wurzelelement.");
            }

            List<XElement> candidates = page.Descendants(One + "T")
                .Where(element =>
                    string.Equals(
                        (string)element.Attribute("selected"),
                        "all",
                        StringComparison.Ordinal) &&
                    element.Parent != null &&
                    element.Parent.Name == One + "OE" &&
                    element.Nodes().Count() == 1 &&
                    element.Nodes().Single() is XCData &&
                    element.Value.Length == 0)
                .ToList();
            if (candidates.Count != 1)
            {
                throw new InvalidOperationException(
                    "OneNote lieferte keinen eindeutigen leeren Cursoranker.");
            }

            XElement caret = candidates[0];
            List<XElement> selectedElements = page.Root
                .DescendantsAndSelf()
                .Where(element => element.Attribute("selected") != null)
                .ToList();
            bool validSelectionHierarchy = selectedElements.All(element =>
                ReferenceEquals(element, caret) ||
                caret.Ancestors().Any(ancestor => ReferenceEquals(ancestor, element)) &&
                string.Equals(
                    (string)element.Attribute("selected"),
                    "partial",
                    StringComparison.Ordinal));
            if (!validSelectionHierarchy)
            {
                throw new InvalidOperationException(
                    "OneNote lieferte neben dem Cursoranker weitere Auswahlmarkierungen.");
            }

            string objectId = caret
                .AncestorsAndSelf()
                .TakeWhile(element => element.Name != One + "Page")
                .Select(element => (string)element.Attribute("objectID"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (string.IsNullOrWhiteSpace(objectId))
            {
                throw new InvalidOperationException(
                    "Der Cursoranker besitzt keine von OneNote vergebene objectID.");
            }
            return objectId;
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
            if (lastText == null)
            {
                throw new InvalidOperationException("Der Textbaustein enthält keinen Textknoten.");
            }
            ClearSelection(lastText.Document);
            lastText.AddAfterSelf(CreateCaretText());
        }

        private static void InsertBlocks(IList<XElement> markers, IList<XElement> incoming)
        {
            List<XElement> paragraphs = markers
                .Select(GetContainingParagraph)
                .Distinct()
                .ToList();
            if (paragraphs.Count == 0 || paragraphs.Any(paragraph => paragraph.Parent == null))
            {
                throw new InvalidOperationException("Die OneNote-Auswahl besitzt keinen eindeutigen Absatzanker.");
            }

            if (paragraphs.Count == 1)
            {
                InsertBlocksWithinParagraph(paragraphs[0], markers, incoming);
                return;
            }

            XElement parent = paragraphs[0].Parent;
            if (paragraphs.Any(paragraph => !ReferenceEquals(paragraph.Parent, parent)))
            {
                throw new InvalidOperationException("Die OneNote-Auswahl überspannt nicht kompatible Container.");
            }

            XElement firstParagraph = paragraphs[0];
            XElement lastParagraph = paragraphs[paragraphs.Count - 1];
            List<XElement> siblings = parent.Elements(One + "OE").ToList();
            int firstParagraphIndex = siblings.IndexOf(firstParagraph);
            int lastParagraphIndex = siblings.IndexOf(lastParagraph);
            if (firstParagraphIndex < 0 || lastParagraphIndex <= firstParagraphIndex)
            {
                throw new InvalidOperationException("Die Reihenfolge der OneNote-Auswahl ist mehrdeutig.");
            }

            RemoveFromFirstSelectionToEnd(firstParagraph,
                markers.Where(marker => ReferenceEquals(GetContainingParagraph(marker), firstParagraph)).ToList());
            RemoveFromStartThroughLastSelection(lastParagraph,
                markers.Where(marker => ReferenceEquals(GetContainingParagraph(marker), lastParagraph)).ToList());

            foreach (XElement sibling in siblings
                .Skip(firstParagraphIndex + 1)
                .Take(lastParagraphIndex - firstParagraphIndex - 1)
                .ToList())
            {
                sibling.Remove();
            }

            bool keepFirst = HasParagraphContent(firstParagraph);
            bool keepLast = HasParagraphContent(lastParagraph);
            if (keepFirst)
            {
                firstParagraph.AddAfterSelf(incoming.Cast<object>().ToArray());
            }
            else
            {
                firstParagraph.AddBeforeSelf(incoming.Cast<object>().ToArray());
                firstParagraph.Remove();
            }

            if (!keepLast) lastParagraph.Remove();
            if (keepLast) SetCaretAtStart(lastParagraph);
            else SetCaretAfter(incoming.Last());
        }

        private static void InsertAfterContainingParagraph(XElement marker, IList<XElement> incoming)
        {
            XElement paragraph = marker.Ancestors(One + "OE").FirstOrDefault();
            if (paragraph == null || paragraph.Parent == null)
            {
                throw new InvalidOperationException("Der Spezialcursor besitzt keinen eindeutigen Absatzanker.");
            }

            paragraph.AddAfterSelf(incoming.Cast<object>().ToArray());
            SetCaretAfter(incoming.Last());
        }

        private static void SetCaretAtStart(XElement paragraph)
        {
            ClearSelection(paragraph.Document);
            XElement caret = CreateCaretText();
            XElement metadata = paragraph.Elements()
                .TakeWhile(IsParagraphMetadata)
                .LastOrDefault();
            if (metadata == null) paragraph.AddFirst(caret);
            else metadata.AddAfterSelf(caret);
        }

        private static void SetCaretAfter(XElement element)
        {
            ClearSelection(element.Document);
            var caret = new XElement(One + "OE",
                CreateCaretText());
            element.AddAfterSelf(caret);
        }

        private static XElement CreateCaretText()
        {
            return new XElement(One + "T",
                new XAttribute("selected", "all"),
                new XCData(string.Empty));
        }

        private static OneNoteInsertionResult CompleteInsertion(
            XDocument page,
            OneNoteSelectionInfo selection,
            string strategy)
        {
            ValidatePostInsertCaret(page);
            return new OneNoteInsertionResult(selection, strategy);
        }

        private static void ClearSelection(XDocument document)
        {
            foreach (XAttribute selected in document.Descendants().Attributes("selected").ToList()) selected.Remove();
        }

        private static SelectionContext CreateSelectionContext(XDocument page)
        {
            List<XElement> rawMarkers = page.Descendants().Where(IsSelected).ToList();
            List<XElement> markers = rawMarkers
                .Where(marker => !marker.Ancestors().Any(IsFullySelected))
                .Where(marker => !(IsPartial(marker) && marker.Descendants().Any(IsSelected)))
                .ToList();
            string markerTypes = rawMarkers.Count == 0
                ? "none"
                : string.Join(",", rawMarkers
                    .GroupBy(marker => marker.Name.LocalName)
                    .OrderBy(group => group.Key, StringComparer.Ordinal)
                    .Select(group => group.Key + ":" + group.Count()));
            string markerPaths = rawMarkers.Count == 0
                ? "none"
                : string.Join(",", rawMarkers.Select(GetElementPath));
            string anchorPath = markers.Count == 0 ? "none" : GetElementPath(markers[0]);
            string scope = "Missing";
            string containerScope = "Page";
            bool usable = false;
            string problem = "OneNote hat die Cursor- oder Auswahlmarkierung verloren. Es wurde nichts eingefügt.";

            if (markers.Count > 0)
            {
                List<XElement> outlines = markers
                    .Select(marker => marker.AncestorsAndSelf(One + "Outline").FirstOrDefault())
                    .Distinct()
                    .ToList();
                List<XElement> cells = markers
                    .Select(marker => marker.AncestorsAndSelf(One + "Cell").FirstOrDefault())
                    .Where(cell => cell != null)
                    .Distinct()
                    .ToList();
                containerScope = cells.Count == 1 && markers.All(marker =>
                    marker.AncestorsAndSelf().Any(ancestor => ReferenceEquals(ancestor, cells[0])))
                    ? "TableCell"
                    : "Outline";

                if (outlines.Count != 1 || outlines[0] == null)
                {
                    scope = "Ambiguous";
                    problem = "Die OneNote-Auswahl überspannt mehrere oder unbekannte Outlines. Es wurde nichts eingefügt.";
                }
                else if (markers.Any(marker => marker.Name == One + "Cell"))
                {
                    if (markers.Count == 1 && markers[0].Name == One + "Cell")
                    {
                        scope = "TableCell";
                        containerScope = "TableCell";
                        usable = true;
                        problem = null;
                    }
                    else
                    {
                        scope = "Ambiguous";
                        problem = "Mehrere Tabellenzellen oder gemischte Zellinhalte sind als Einfügeanker mehrdeutig.";
                    }
                }
                else if (markers.Any(marker => marker.Name != One + "T" && IsPartial(marker)))
                {
                    scope = "Ambiguous";
                    problem = "OneNote meldet nur einen unvollständigen Blockanker ohne präzise Nachfahrmarkierung.";
                }
                else if (markers.All(marker => marker.Name == One + "T"))
                {
                    bool allEmpty = markers.All(marker => marker.Value.Length == 0);
                    if (allEmpty && markers.Count > 1)
                    {
                        scope = "Ambiguous";
                        problem = "OneNote meldet mehrere leere Textcursor. Es wurde nichts eingefügt.";
                    }
                    else if (markers.Any(marker => GetContainingParagraph(marker) == null))
                    {
                        scope = "Ambiguous";
                        problem = "Die Textauswahl besitzt keinen eindeutigen Absatzanker.";
                    }
                    else if (markers.Count == 1 && ContainsSpecialMarkup(markers[0]))
                    {
                        scope = "SpecialCursor";
                        usable = true;
                        problem = null;
                    }
                    else
                    {
                        scope = allEmpty ? "TextCursor" : "TextRange";
                        usable = HasCompatibleParagraphRange(markers);
                        problem = usable
                            ? null
                            : "Die Textauswahl überspannt nicht kompatible Absatzcontainer.";
                    }
                }
                else if (markers.All(marker => GetContainingParagraph(marker) != null))
                {
                    scope = "Block";
                    usable = HasCompatibleParagraphRange(markers);
                    problem = usable
                        ? null
                        : "Die Blockauswahl überspannt nicht kompatible Absatzcontainer.";
                }
                else
                {
                    scope = "Ambiguous";
                    problem = "Die OneNote-Auswahl besitzt keinen unterstützten Einfügeanker.";
                }
            }

            return new SelectionContext
            {
                Markers = markers,
                Info = new OneNoteSelectionInfo(
                    scope,
                    containerScope,
                    rawMarkers.Count,
                    markerTypes,
                    markerPaths,
                    anchorPath,
                    usable,
                    problem)
            };
        }

        private static bool HasCompatibleParagraphRange(IList<XElement> markers)
        {
            List<XElement> paragraphs = markers
                .Select(GetContainingParagraph)
                .Distinct()
                .ToList();
            if (paragraphs.Count <= 1) return true;
            XElement parent = paragraphs[0].Parent;
            return parent != null &&
                paragraphs.All(paragraph => ReferenceEquals(paragraph.Parent, parent)) &&
                parent.Elements(One + "OE").ToList().IndexOf(paragraphs[0]) <
                parent.Elements(One + "OE").ToList().IndexOf(paragraphs[paragraphs.Count - 1]);
        }

        private static XElement GetContainingParagraph(XElement marker)
        {
            if (marker == null) return null;
            return marker.Name == One + "OE"
                ? marker
                : marker.Ancestors(One + "OE").FirstOrDefault();
        }

        private static string GetElementPath(XElement element)
        {
            var parts = new Stack<string>();
            for (XElement current = element; current != null; current = current.Parent)
            {
                int index = current.ElementsBeforeSelf(current.Name).Count() + 1;
                parts.Push(current.Name.LocalName + "[" + index + "]");
            }
            return "/" + string.Join("/", parts);
        }

        private static void InsertIntoTableCell(XElement cell, IList<XElement> incoming)
        {
            XElement children = cell.Element(One + "OEChildren");
            if (children == null)
            {
                children = new XElement(One + "OEChildren");
                cell.Add(children);
            }
            else
            {
                children.RemoveNodes();
            }
            children.Add(incoming.Cast<object>().ToArray());
            SetCaretAfter(incoming.Last());
        }

        private static void InsertBlocksWithinParagraph(
            XElement paragraph,
            IList<XElement> markers,
            IList<XElement> incoming)
        {
            if (markers.Any(marker => ReferenceEquals(marker, paragraph)))
            {
                paragraph.AddBeforeSelf(incoming.Cast<object>().ToArray());
                paragraph.Remove();
                SetCaretAfter(incoming.Last());
                return;
            }

            List<XElement> direct = paragraph.Elements().ToList();
            List<XElement> selectedChildren = markers
                .Select(marker => GetDirectParagraphChild(paragraph, marker))
                .Where(child => child != null)
                .ToList();
            if (selectedChildren.Count == 0)
            {
                throw new InvalidOperationException("Die Blockauswahl besitzt keinen direkten Absatzanker.");
            }

            int firstIndex = selectedChildren.Select(child => direct.IndexOf(child)).Min();
            int lastIndex = selectedChildren.Select(child => direct.IndexOf(child)).Max();
            if (firstIndex < 0 || lastIndex < firstIndex)
            {
                throw new InvalidOperationException("Die Blockauswahl besitzt eine mehrdeutige Reihenfolge.");
            }

            var suffix = new XElement(One + "OE");
            CopyFormattingAttributes(paragraph, suffix);
            foreach (XElement metadata in direct.Take(firstIndex).Where(IsParagraphMetadata))
            {
                suffix.Add(new XElement(metadata));
            }
            foreach (XElement element in direct.Skip(lastIndex + 1).ToList())
            {
                element.Remove();
                suffix.Add(element);
            }
            foreach (XElement element in direct.Skip(firstIndex).Take(lastIndex - firstIndex + 1).ToList())
            {
                if (element.Parent != null) element.Remove();
            }

            bool keepPrefix = HasParagraphContent(paragraph);
            if (keepPrefix)
            {
                paragraph.AddAfterSelf(incoming.Cast<object>().ToArray());
            }
            else
            {
                paragraph.AddBeforeSelf(incoming.Cast<object>().ToArray());
                paragraph.Remove();
            }

            if (HasParagraphContent(suffix))
            {
                incoming.Last().AddAfterSelf(suffix);
                SetCaretAtStart(suffix);
            }
            else
            {
                SetCaretAfter(incoming.Last());
            }
        }

        private static XElement GetDirectParagraphChild(XElement paragraph, XElement marker)
        {
            XElement current = marker;
            while (current != null && current.Parent != null && !ReferenceEquals(current.Parent, paragraph))
            {
                current = current.Parent;
            }
            return current != null && ReferenceEquals(current.Parent, paragraph) ? current : null;
        }

        private static void RemoveFromFirstSelectionToEnd(
            XElement paragraph,
            IList<XElement> markers)
        {
            if (markers.Any(marker => ReferenceEquals(marker, paragraph)))
            {
                paragraph.RemoveNodes();
                return;
            }
            List<XElement> direct = paragraph.Elements().ToList();
            int first = markers
                .Select(marker => GetDirectParagraphChild(paragraph, marker))
                .Where(child => child != null)
                .Select(child => direct.IndexOf(child))
                .DefaultIfEmpty(-1)
                .Min();
            if (first < 0) throw new InvalidOperationException("Der Start der Blockauswahl ist mehrdeutig.");
            foreach (XElement element in direct.Skip(first).ToList()) element.Remove();
        }

        private static void RemoveFromStartThroughLastSelection(
            XElement paragraph,
            IList<XElement> markers)
        {
            if (markers.Any(marker => ReferenceEquals(marker, paragraph)))
            {
                paragraph.RemoveNodes();
                return;
            }
            List<XElement> direct = paragraph.Elements().ToList();
            int last = markers
                .Select(marker => GetDirectParagraphChild(paragraph, marker))
                .Where(child => child != null)
                .Select(child => direct.IndexOf(child))
                .DefaultIfEmpty(-1)
                .Max();
            if (last < 0) throw new InvalidOperationException("Das Ende der Blockauswahl ist mehrdeutig.");
            foreach (XElement element in direct.Take(last + 1).ToList())
            {
                bool precedingParagraphMetadata = IsParagraphMetadata(element) &&
                    direct.IndexOf(element) < last &&
                    !markers.Any(marker => ReferenceEquals(marker, element));
                if (!precedingParagraphMetadata) element.Remove();
            }
        }

        private static bool HasParagraphContent(XElement paragraph)
        {
            return paragraph.Elements().Any(element =>
                element.Name != One + "List" && element.Name != One + "Tag");
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
            return !incoming[0].Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) &&
                incoming[0].Elements(One + "T").Any() &&
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
                    throw new InvalidOperationException("Ein Bild enthält weder Data noch callbackID.");
                }
                if (binaryResolver == null) throw new InvalidOperationException("Für ein Callback-Bild ist kein Binärdatenleser verfügbar.");
                string data = binaryResolver(callback);
                if (string.IsNullOrWhiteSpace(data)) throw new InvalidOperationException("OneNote lieferte keine Binärdaten für Bild " + callback + ".");
                image.Add(new XElement(One + "Data", data));
                XAttribute callbackAttribute = image.Attribute("callbackID");
                if (callbackAttribute != null) callbackAttribute.Remove();
            }
        }

        private static void PruneUnselectedContent(XElement element, bool selectedAncestor)
        {
            bool keepWholeSubtree = selectedAncestor || IsFullySelected(element);
            if (keepWholeSubtree) return;

            foreach (XElement child in element.Elements().ToList())
            {
                bool containsSelection = IsSelected(child) || child.Descendants().Any(IsSelected);
                bool structuralMetadata =
                    element.Name == One + "OE" && IsParagraphMetadata(child) ||
                    element.Name == One + "Table" && child.Name != One + "Row";
                if (!containsSelection && !structuralMetadata)
                {
                    if (element.Name == One + "Row" && child.Name == One + "Cell")
                    {
                        child.ReplaceWith(CreateEmptyTableCell(child));
                        continue;
                    }
                    child.Remove();
                    continue;
                }
                if (containsSelection) PruneUnselectedContent(child, false);
            }
        }

        private static bool IsParagraphMetadata(XElement element)
        {
            return element.Name == One + "List" || element.Name == One + "Tag";
        }

        private static XElement CreateEmptyTableCell(XElement source)
        {
            var cell = new XElement(One + "Cell");
            foreach (XAttribute attribute in source.Attributes().Where(attribute =>
                attribute.Name.LocalName != "selected" &&
                attribute.Name.LocalName != "objectID" &&
                attribute.Name.LocalName != "lastModifiedTime"))
            {
                cell.Add(new XAttribute(attribute));
            }
            cell.Add(
                new XElement(One + "OEChildren",
                    new XElement(One + "OE",
                        new XElement(One + "T", new XCData(string.Empty)))));
            return cell;
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

        private static bool IsFullySelected(XElement element)
        {
            string value = (string)element.Attribute("selected");
            return string.Equals(value, "all", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsPartial(XElement element)
        {
            return string.Equals(
                (string)element.Attribute("selected"),
                "partial",
                StringComparison.OrdinalIgnoreCase);
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
            throw new InvalidOperationException("Nicht unterstützter XML-Knotentyp: " + node.NodeType);
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
