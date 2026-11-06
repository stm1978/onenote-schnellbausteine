using System;
using System.Collections.Generic;
using System.Globalization;
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
            public XElement Outline;
        }

        private sealed class CaptureContext
        {
            public List<XElement> Roots;
            public XElement Outline;
            public bool PruneRoots;
        }

        private sealed class MixedCaptureItem
        {
            public XElement Source;
            public XElement Root;
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
            if (!HasSelection(page)) throw new InvalidOperationException("OneNote meldet keine markierte Auswahl.");
            CaptureContext capture = CreateCaptureContext(page);

            var content = new XElement(Snippet + "Content");
            foreach (XElement root in capture.Roots)
            {
                XElement clone = new XElement(root);
                if (capture.PruneRoots)
                {
                    PruneUnselectedContent(clone, false);
                }
                RemoveTransientAttributes(clone);
                ResolveBinaryData(clone, binaryResolver);
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

            var snippet = new XElement(Snippet + "Snippet",
                new XAttribute("version", "1"));
            var geometry = new XElement(Snippet + "Geometry");
            double minimumWidth;
            if (TryDeriveMinimumWidth(capture.Outline, content, out minimumWidth))
            {
                geometry.SetAttributeValue("minWidth", FormatWidth(minimumWidth));
            }
            bool relativeInkZ;
            if (NormalizeInkPositions(content, out relativeInkZ))
            {
                geometry.SetAttributeValue("inkPositions", "relative");
                if (relativeInkZ) geometry.SetAttributeValue("inkPositionZ", "relative");
            }
            if (geometry.HasAttributes) snippet.Add(geometry);
            snippet.Add(definitions, quickStyles, content);
            return new XDocument(snippet);
        }

        private static CaptureContext CreateCaptureContext(XDocument page)
        {
            List<XElement> rawMarkers = page.Descendants().Where(IsSelected).ToList();
            List<XElement> selectedInk = page.Descendants(One + "InkDrawing")
                .Where(ink =>
                    IsSelected(ink) ||
                    ink.Descendants().Any(IsSelected) ||
                    ink.Ancestors().Any(IsFullySelected))
                .ToList();
            if (selectedInk.Count > 0)
            {
                List<XElement> effectiveMarkers = GetEffectiveCaptureMarkers(rawMarkers);
                bool containsOtherContent = effectiveMarkers.Any(marker =>
                    ContainsContentOutsideInk(marker, selectedInk));
                return containsOtherContent
                    ? CreateMixedCaptureContext(rawMarkers, effectiveMarkers, selectedInk)
                    : CreateInkCaptureContext(rawMarkers, selectedInk);
            }

            List<XElement> structuralTables = FindSelectedTables(rawMarkers);
            if (structuralTables.Count > 0)
            {
                return CreateTableCaptureContext(rawMarkers, structuralTables);
            }

            return CreateContentCaptureContext(rawMarkers);
        }

        private static CaptureContext CreateInkCaptureContext(
            IList<XElement> rawMarkers,
            IList<XElement> selectedInk)
        {
            ValidateSelectedInk(selectedInk);
            if (rawMarkers.Any(marker =>
                !selectedInk.Any(ink => IsRelatedSelectionMarker(marker, ink))))
            {
                throw new InvalidOperationException(
                    "Freihand und andere Inhalte können nicht gemeinsam als Schnellbaustein gespeichert werden.");
            }

            List<XElement> orderedInk = selectedInk
                .OrderBy(ink => ink, XNode.DocumentOrderComparer)
                .ToList();
            return new CaptureContext
            {
                Roots = orderedInk.Select(CreateInkRoot).ToList(),
                Outline = GetSingleOutlineOrNull(orderedInk),
                PruneRoots = false
            };
        }

        private static CaptureContext CreateMixedCaptureContext(
            IList<XElement> rawMarkers,
            IList<XElement> effectiveMarkers,
            IList<XElement> selectedInk)
        {
            ValidateSelectedInk(selectedInk);

            List<XElement> structuralTables = FindSelectedTables(rawMarkers);
            if (structuralTables.Count > 1)
            {
                throw new InvalidOperationException(
                    "Gemischte Auswahlen mit mehreren Tabellen sind nicht eindeutig genug.");
            }

            XElement table = structuralTables.SingleOrDefault();
            if (table != null) ValidateTableSelectionMarkers(table, rawMarkers);

            List<XElement> contentMarkers = effectiveMarkers
                .Where(marker => ContainsContentOutsideInk(marker, selectedInk))
                .Where(marker => table == null || !IsWithin(marker, table))
                .ToList();
            if (contentMarkers.Any(marker =>
                marker.Name != One + "T" && IsPartial(marker)))
            {
                throw new InvalidOperationException(
                    "Gemischte Inhalte enthalten einen unvollständigen Block ohne präzise Nachfahrmarkierung.");
            }
            List<XElement> contentRoots = contentMarkers.Count == 0
                ? new List<XElement>()
                : BuildContentCaptureRoots(contentMarkers);
            if (contentRoots.Any(root =>
                selectedInk.Any(ink => IsWithin(root, ink)) ||
                table != null && IsWithin(table, root)))
            {
                throw new InvalidOperationException(
                    "Freihand und andere Inhalte besitzen überlappende Auswahlwurzeln.");
            }

            var items = new List<MixedCaptureItem>();
            foreach (XElement root in contentRoots)
            {
                XElement clone = new XElement(root);
                bool selectedAncestor = root.Ancestors().Any(IsFullySelected);
                if (selectedAncestor) clone.SetAttributeValue("selected", "all");
                RemoveSelectedInk(clone);
                PruneUnselectedContent(clone, selectedAncestor);
                if (!clone.DescendantsAndSelf().Any(IsSelected) ||
                    !clone.DescendantsAndSelf().Any(IsContentElement))
                {
                    throw new InvalidOperationException(
                        "Ein gemischter Inhaltsblock enthält keinen eindeutig ausgewählten Inhalt.");
                }
                XElement source = contentMarkers.FirstOrDefault(marker =>
                    IsWithin(marker, root) || IsWithin(root, marker)) ?? root;
                items.Add(new MixedCaptureItem { Source = source, Root = clone });
            }
            if (table != null)
            {
                items.Add(new MixedCaptureItem
                {
                    Source = table,
                    Root = new XElement(One + "OE", ExtractSelectedTableRectangle(table))
                });
            }
            foreach (XElement ink in selectedInk)
            {
                items.Add(new MixedCaptureItem { Source = ink, Root = CreateInkRoot(ink) });
            }

            if (!items.Any(item => item.Root.DescendantsAndSelf(One + "InkDrawing").Any()) ||
                !items.Any(item => !item.Root.DescendantsAndSelf(One + "InkDrawing").Any()))
            {
                throw new InvalidOperationException(
                    "Die gemischte Auswahl enthält keine eindeutigen Ink- und Inhaltswurzeln.");
            }

            List<MixedCaptureItem> ordered = items
                .OrderBy(item => item.Source, XNode.DocumentOrderComparer)
                .ToList();
            return new CaptureContext
            {
                Roots = ordered.Select(item => item.Root).ToList(),
                Outline = GetSingleOutlineOrNull(ordered.Select(item => item.Source)),
                PruneRoots = false
            };
        }

        private static void ValidateSelectedInk(IList<XElement> selectedInk)
        {
            if (selectedInk.Count == 0 ||
                selectedInk.Any(ink =>
                    !IsFullySelected(ink) &&
                    !ink.Ancestors().Any(IsFullySelected)))
            {
                throw new InvalidOperationException(
                    "Die Freihandauswahl enthält keinen vollständig gelieferten InkDrawing-Binärblob.");
            }
            if (selectedInk.Any(ink => ink.Ancestors().Any(ancestor =>
                IsFullySelected(ancestor) && ancestor.Name != One + "OE")))
            {
                throw new InvalidOperationException(
                    "Eine vollständig markierte übergeordnete Seitenstruktur ist für Freihand nicht eindeutig.");
            }
            if (selectedInk.Any(ink => selectedInk.Any(other =>
                !ReferenceEquals(ink, other) && IsWithin(ink, other))))
            {
                throw new InvalidOperationException(
                    "Die Freihandauswahl besitzt überlappende InkDrawing-Wurzeln.");
            }
        }

        private static List<XElement> FindSelectedTables(
            IEnumerable<XElement> rawMarkers)
        {
            return rawMarkers
                .SelectMany(marker => IsFullySelected(marker)
                    ? marker.DescendantsAndSelf(One + "Table")
                    : marker.AncestorsAndSelf(One + "Table"))
                .Where(table =>
                    table.AncestorsAndSelf().Any(IsFullySelected) ||
                    table.Descendants(One + "Row").Any(IsFullySelected) ||
                    table.Descendants(One + "Cell").Any(IsFullySelected))
                .Distinct()
                .ToList();
        }

        private static XElement CreateInkRoot(XElement ink)
        {
            XElement clone = new XElement(ink);
            return new XElement(One + "OE", clone);
        }

        private static CaptureContext CreateTableCaptureContext(
            IList<XElement> rawMarkers,
            IList<XElement> structuralTables)
        {
            if (structuralTables.Count != 1)
            {
                throw new InvalidOperationException(
                    "Eine Auswahl aus mehreren Tabellen kann nicht als Schnellbaustein gespeichert werden.");
            }

            XElement table = structuralTables[0];
            if (rawMarkers.Any(marker => !IsRelatedSelectionMarker(marker, table)))
            {
                throw new InvalidOperationException(
                    "Tabellen und externe Inhalte können nicht gemeinsam als Schnellbaustein gespeichert werden.");
            }
            ValidateTableSelectionMarkers(table, rawMarkers);
            foreach (XElement ancestor in table.Ancestors().Where(IsFullySelected))
            {
                if (ancestor.Descendants().Any(element =>
                    !ReferenceEquals(element, table) &&
                    element.Name != One + "Table" &&
                    !element.Ancestors().Any(parent => ReferenceEquals(parent, table)) &&
                    IsContentElement(element)))
                {
                    throw new InvalidOperationException(
                        "Die Auswahl enthält neben der Tabelle weitere Inhalte.");
                }
            }

            XElement compactTable = ExtractSelectedTableRectangle(table);
            XElement outline = table.Ancestors(One + "Outline").FirstOrDefault();
            return new CaptureContext
            {
                Roots = new List<XElement> { new XElement(One + "OE", compactTable) },
                Outline = outline,
                PruneRoots = false
            };
        }

        private static void ValidateTableSelectionMarkers(
            XElement table,
            IList<XElement> rawMarkers)
        {
            List<XElement> selectedCells = table.Descendants(One + "Cell")
                .Where(cell =>
                    IsFullySelected(cell) ||
                    cell.Ancestors().Any(ancestor =>
                        ReferenceEquals(ancestor, table) && IsFullySelected(ancestor) ||
                        ancestor.Name == One + "Row" && IsFullySelected(ancestor)))
                .ToList();
            if (rawMarkers.Any(marker =>
                marker.Ancestors().Any(ancestor => ReferenceEquals(ancestor, table)) &&
                !selectedCells.Any(cell => IsRelatedSelectionMarker(marker, cell))))
            {
                throw new InvalidOperationException(
                    "Die Tabellenmarkierung enthält zusätzliche Teilinhalte außerhalb des vollständigen Zellbereichs.");
            }
        }

        private static XElement ExtractSelectedTableRectangle(XElement table)
        {
            if (IsFullySelected(table)) return new XElement(table);

            List<XElement> rows = table.Elements(One + "Row").ToList();
            if (rows.Count == 0)
            {
                throw new InvalidOperationException("Die ausgewählte Tabelle enthält keine Zeilen.");
            }
            List<List<XElement>> cells = rows
                .Select(row => row.Elements(One + "Cell").ToList())
                .ToList();
            int columnCount = cells[0].Count;
            if (columnCount == 0 || cells.Any(row => row.Count != columnCount))
            {
                throw new InvalidOperationException(
                    "Die ausgewählte Tabelle besitzt keine eindeutige rechteckige Zellstruktur.");
            }

            var selectedCoordinates = new List<Tuple<int, int>>();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                bool wholeRow = IsFullySelected(rows[rowIndex]) || IsFullySelected(table);
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    if (wholeRow || IsFullySelected(cells[rowIndex][columnIndex]))
                    {
                        selectedCoordinates.Add(Tuple.Create(rowIndex, columnIndex));
                    }
                }
            }
            if (selectedCoordinates.Count == 0)
            {
                throw new InvalidOperationException(
                    "Die Tabellenmarkierung enthält keine vollständig ausgewählten Zellen.");
            }

            int firstRow = selectedCoordinates.Min(item => item.Item1);
            int lastRow = selectedCoordinates.Max(item => item.Item1);
            int firstColumn = selectedCoordinates.Min(item => item.Item2);
            int lastColumn = selectedCoordinates.Max(item => item.Item2);
            int expectedCount = (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
            if (selectedCoordinates.Count != expectedCount)
            {
                throw new InvalidOperationException(
                    "Nicht rechteckige Tabellenbereiche können nicht als Schnellbaustein gespeichert werden.");
            }

            var result = new XElement(table.Name);
            foreach (XAttribute attribute in table.Attributes().Where(attribute =>
                !attribute.IsNamespaceDeclaration &&
                attribute.Name.LocalName != "selected"))
            {
                result.Add(new XAttribute(attribute));
            }

            XElement sourceColumns = table.Element(One + "Columns");
            if (sourceColumns != null)
            {
                var columns = new XElement(sourceColumns.Name);
                foreach (XAttribute attribute in sourceColumns.Attributes())
                {
                    columns.Add(new XAttribute(attribute));
                }
                List<XElement> sourceColumnList = sourceColumns.Elements(One + "Column").ToList();
                if (sourceColumnList.Count != columnCount)
                {
                    throw new InvalidOperationException(
                        "Die Spaltendefinitionen stimmen nicht mit der Zellstruktur überein.");
                }
                for (int columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++)
                {
                    XElement column = new XElement(sourceColumnList[columnIndex]);
                    column.SetAttributeValue("index", columnIndex - firstColumn);
                    columns.Add(column);
                }
                result.Add(columns);
            }

            foreach (XElement metadata in table.Elements().Where(element =>
                element.Name != One + "Columns" && element.Name != One + "Row"))
            {
                result.Add(new XElement(metadata));
            }
            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
            {
                XElement row = new XElement(rows[rowIndex]);
                row.Elements(One + "Cell").Remove();
                foreach (XElement cell in cells[rowIndex]
                    .Skip(firstColumn)
                    .Take(lastColumn - firstColumn + 1))
                {
                    row.Add(new XElement(cell));
                }
                result.Add(row);
            }
            return result;
        }

        private static CaptureContext CreateContentCaptureContext(IList<XElement> rawMarkers)
        {
            List<XElement> markers = GetEffectiveCaptureMarkers(rawMarkers);
            if (markers.Any(marker => marker.Name != One + "T" && IsPartial(marker)))
            {
                throw new InvalidOperationException(
                    "OneNote meldet nur einen unvollständigen Block ohne präzise Nachfahrmarkierung.");
            }
            List<XElement> roots = BuildContentCaptureRoots(markers);
            List<XElement> outlines = roots
                .Select(root => root.Ancestors(One + "Outline").FirstOrDefault())
                .Distinct()
                .ToList();
            if (outlines.Count != 1 || outlines[0] == null)
            {
                throw new InvalidOperationException(
                    "Die Auswahl überspannt mehrere oder unbekannte Outlines.");
            }
            return new CaptureContext
            {
                Roots = roots,
                Outline = outlines[0],
                PruneRoots = true
            };
        }

        private static List<XElement> GetEffectiveCaptureMarkers(
            IEnumerable<XElement> rawMarkers)
        {
            return rawMarkers
                .Where(marker => !marker.Ancestors().Any(IsFullySelected))
                .Where(marker => !(IsPartial(marker) && marker.Descendants().Any(IsSelected)))
                .ToList();
        }

        private static List<XElement> BuildContentCaptureRoots(
            IList<XElement> markers)
        {
            var roots = new List<XElement>();
            foreach (XElement marker in markers)
            {
                XElement candidate = marker.Name == One + "OE"
                    ? marker
                    : marker.Ancestors(One + "OE").FirstOrDefault();
                if (candidate == null)
                {
                    if (marker.Name == One + "Title" ||
                        marker.Name == One + "Outline" ||
                        marker.Name == One + "OEChildren")
                    {
                        int rootCount = roots.Count;
                        foreach (XElement outlineElement in marker.Descendants(One + "OE"))
                        {
                            AddCaptureRoot(roots, outlineElement);
                        }
                        if (roots.Count > rootCount) continue;
                    }

                    throw new InvalidOperationException(
                        "Die Auswahl enthält einen Inhalt ohne eindeutige OE-Wurzel.");
                }

                XElement richAncestor = marker.Ancestors().FirstOrDefault(element =>
                    element.Name == One + "Table" || element.Name == One + "Image");
                if (richAncestor != null)
                {
                    candidate = richAncestor.Ancestors(One + "OE").FirstOrDefault() ?? candidate;
                }
                AddCaptureRoot(roots, candidate);
            }
            if (roots.Count == 0)
            {
                throw new InvalidOperationException(
                    "Die OneNote-Auswahl enthält keinen speicherbaren Inhalt.");
            }
            return roots;
        }

        private static bool ContainsContentOutsideInk(
            XElement marker,
            IList<XElement> selectedInk)
        {
            return marker.DescendantsAndSelf().Any(element =>
                IsContentElement(element) &&
                !selectedInk.Any(ink => IsWithin(element, ink)));
        }

        private static void RemoveSelectedInk(XElement content)
        {
            foreach (XElement ink in content.DescendantsAndSelf(One + "InkDrawing")
                .Where(element =>
                    IsSelected(element) ||
                    element.Ancestors().Any(IsFullySelected))
                .ToList())
            {
                ink.Remove();
            }
        }

        private static void AddCaptureRoot(ICollection<XElement> roots, XElement candidate)
        {
            if (roots.Any(root => ReferenceEquals(root, candidate)) ||
                roots.Any(root => candidate.Ancestors().Any(ancestor => ReferenceEquals(ancestor, root))))
            {
                return;
            }
            foreach (XElement root in roots
                .Where(root => root.Ancestors().Any(ancestor => ReferenceEquals(ancestor, candidate)))
                .ToList())
            {
                roots.Remove(root);
            }
            roots.Add(candidate);
        }

        private static bool IsRelatedSelectionMarker(XElement marker, XElement content)
        {
            return ReferenceEquals(marker, content) ||
                marker.Ancestors().Any(ancestor => ReferenceEquals(ancestor, content)) ||
                content.Ancestors().Any(ancestor => ReferenceEquals(ancestor, marker));
        }

        private static bool IsWithin(XElement element, XElement root)
        {
            return ReferenceEquals(element, root) ||
                element.Ancestors().Any(ancestor => ReferenceEquals(ancestor, root));
        }

        private static XElement GetSingleOutlineOrNull(IEnumerable<XElement> elements)
        {
            List<XElement> outlines = elements
                .Select(element => element.AncestorsAndSelf(One + "Outline").FirstOrDefault())
                .Where(outline => outline != null)
                .Distinct()
                .ToList();
            return outlines.Count == 1 ? outlines[0] : null;
        }

        private static bool IsContentElement(XElement element)
        {
            return element.Name == One + "T" ||
                element.Name == One + "Image" ||
                element.Name == One + "InkDrawing" ||
                element.Name == One + "InsertedFile";
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
            bool containsInk = incoming
                .SelectMany(element => element.DescendantsAndSelf(One + "InkDrawing"))
                .Any();
            if (containsInk)
            {
                bool validInkFragment =
                    incoming.All(element => element.Name == One + "OE") &&
                    incoming.Where(element =>
                        element.DescendantsAndSelf(One + "InkDrawing").Any())
                        .All(element =>
                            element.Elements().Count() == 1 &&
                            element.Elements(One + "InkDrawing").Count() == 1);
                bool validInkAnchor =
                    selection.Info.Scope == "TableCell" ||
                    selection.Info.Scope == "TextCursor" ||
                    selection.Info.Scope == "TextRange" ||
                    selection.Info.Scope == "Block";
                if (!validInkFragment || !validInkAnchor)
                {
                    throw new InvalidOperationException(
                        "Freihandzeichnungen können nur als eindeutige OE-Wurzeln an einem Text- oder Zellanker eingefügt werden.");
                }
            }

            foreach (XElement element in incoming) RemoveTransientAttributes(element);
            RebaseInkPositions(incoming, selection.Outline, fragment);
            MergeTagDefinitions(page, fragment, incoming);
            MergeQuickStyleDefinitions(page, fragment, incoming);
            ClearSelection(page);

            List<XElement> markers = selection.Markers;
            if (selection.Info.Scope == "TableCell")
            {
                InsertIntoTableCell(markers[0], incoming);
                return CompleteInsertion(
                    page,
                    fragment,
                    selection.Outline,
                    selection.Info,
                    "ReplaceTableCell");
            }

            if (selection.Info.Scope == "SpecialCursor")
            {
                InsertAfterContainingParagraph(markers[0], incoming);
                return CompleteInsertion(
                    page,
                    fragment,
                    selection.Outline,
                    selection.Info,
                    "AfterSpecialParagraph");
            }

            if (CanInsertInline(incoming) &&
                markers.All(marker => marker.Name == One + "T") &&
                markers.Select(marker => marker.Parent).Distinct().Count() == 1)
            {
                InsertInline(markers, incoming[0]);
                return CompleteInsertion(
                    page,
                    fragment,
                    selection.Outline,
                    selection.Info,
                    "Inline");
            }

            InsertBlocks(markers, incoming);
            return CompleteInsertion(
                page,
                fragment,
                selection.Outline,
                selection.Info,
                "BlockRange");
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
            XDocument fragment,
            XElement targetOutline,
            OneNoteSelectionInfo selection,
            string strategy)
        {
            ValidatePostInsertCaret(page);
            ExpandOutlineToMinimumWidth(targetOutline, fragment);
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

            XElement selectedOutline = null;
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

                if (outlines.Count == 1) selectedOutline = outlines[0];
            }

            return new SelectionContext
            {
                Markers = markers,
                Outline = selectedOutline,
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

        private static bool TryDeriveMinimumWidth(
            XElement outline,
            XElement content,
            out double minimumWidth)
        {
            minimumWidth = 0;
            double width;
            if (outline != null &&
                TryReadWidth(outline.Element(One + "Size"), out width))
            {
                minimumWidth = width;
            }

            foreach (XElement table in content.DescendantsAndSelf(One + "Table"))
            {
                if (TryReadTableWidth(table, out width))
                {
                    minimumWidth = Math.Max(minimumWidth, width);
                }
            }
            foreach (XElement image in content.DescendantsAndSelf(One + "Image"))
            {
                if (TryReadWidth(image.Element(One + "Size"), out width))
                {
                    minimumWidth = Math.Max(minimumWidth, width);
                }
            }
            foreach (XElement ink in content.DescendantsAndSelf(One + "InkDrawing"))
            {
                if (TryReadWidth(ink.Element(One + "Size"), out width))
                {
                    minimumWidth = Math.Max(minimumWidth, width);
                }
            }
            return minimumWidth > 0;
        }

        private static bool TryReadTableWidth(XElement table, out double width)
        {
            width = 0;
            XElement columns = table.Element(One + "Columns");
            if (columns == null) return false;
            List<XElement> columnList = columns.Elements(One + "Column").ToList();
            if (columnList.Count == 0) return false;

            foreach (XElement column in columnList)
            {
                double columnWidth;
                if (!TryReadPositiveDouble(
                    (string)column.Attribute("width"),
                    out columnWidth))
                {
                    width = 0;
                    return false;
                }
                width += columnWidth;
            }
            return IsPositiveFinite(width);
        }

        private static bool TryReadWidth(XElement size, out double width)
        {
            return TryReadPositiveDouble(
                size == null ? null : (string)size.Attribute("width"),
                out width);
        }

        private static bool TryReadPositiveDouble(string value, out double result)
        {
            return double.TryParse(
                    value,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out result) &&
                IsPositiveFinite(result);
        }

        private static bool NormalizeInkPositions(XElement content, out bool hasZ)
        {
            hasZ = false;
            List<XElement> inkDrawings = content.DescendantsAndSelf(One + "InkDrawing").ToList();
            if (inkDrawings.Count == 0) return false;

            var positions = new List<Tuple<XElement, double, double, double>>();
            bool missingPosition = false;
            foreach (XElement ink in inkDrawings)
            {
                XElement position = ink.Element(One + "Position");
                double x;
                double y;
                double z = 0;
                if (position == null)
                {
                    missingPosition = true;
                    continue;
                }
                if (!TryReadCoordinate((string)position.Attribute("x"), out x) ||
                    !TryReadCoordinate((string)position.Attribute("y"), out y))
                {
                    throw new InvalidOperationException(
                        "Die Freihandauswahl enthält eine ungültige InkDrawing-Position.");
                }
                string zValue = (string)position.Attribute("z");
                bool currentHasZ = zValue != null;
                if (positions.Count > 0 && currentHasZ != hasZ)
                {
                    throw new InvalidOperationException(
                        "Die Freihandauswahl enthält unvollständige InkDrawing-Positionen.");
                }
                if (positions.Count == 0) hasZ = currentHasZ;
                if (currentHasZ && !TryReadCoordinate(zValue, out z))
                {
                    throw new InvalidOperationException(
                        "Die Freihandauswahl enthält eine ungültige InkDrawing-Position.");
                }
                positions.Add(Tuple.Create(position, x, y, z));
            }

            if (positions.Count == 0) return false;
            if (missingPosition)
            {
                throw new InvalidOperationException(
                    "Die Freihandauswahl enthält InkDrawing-Objekte ohne eindeutige Position.");
            }
            if (positions.Count != inkDrawings.Count)
            {
                throw new InvalidOperationException(
                    "Die Freihandauswahl enthält InkDrawing-Objekte ohne eindeutige Position.");
            }

            double originX = positions.Min(position => position.Item2);
            double originY = positions.Min(position => position.Item3);
            double originZ = hasZ ? positions.Min(position => position.Item4) : 0;
            foreach (Tuple<XElement, double, double, double> position in positions)
            {
                position.Item1.SetAttributeValue("x", FormatWidth(position.Item2 - originX));
                position.Item1.SetAttributeValue("y", FormatWidth(position.Item3 - originY));
                if (hasZ)
                {
                    position.Item1.SetAttributeValue("z", FormatWidth(position.Item4 - originZ));
                }
            }
            return true;
        }

        private static void RebaseInkPositions(
            IList<XElement> roots,
            XElement targetOutline,
            XDocument fragment)
        {
            XElement geometry = fragment.Root.Element(Snippet + "Geometry");
            if (geometry == null ||
                !string.Equals(
                    (string)geometry.Attribute("inkPositions"),
                    "relative",
                    StringComparison.Ordinal))
            {
                return;
            }

            XElement targetPosition = targetOutline == null
                ? null
                : targetOutline.Element(One + "Position");
            double targetX;
            double targetY;
            double targetZ = 0;
            if (targetPosition == null ||
                !TryReadCoordinate((string)targetPosition.Attribute("x"), out targetX) ||
                !TryReadCoordinate((string)targetPosition.Attribute("y"), out targetY))
            {
                throw new InvalidOperationException(
                    "Die Einfügeposition für die Freihandauswahl konnte nicht ermittelt werden.");
            }
            bool rebaseZ = string.Equals(
                (string)geometry.Attribute("inkPositionZ"),
                "relative",
                StringComparison.Ordinal);
            if (rebaseZ &&
                !TryReadCoordinate((string)targetPosition.Attribute("z"), out targetZ))
            {
                throw new InvalidOperationException(
                    "Die räumliche Einfügeposition für die Freihandauswahl konnte nicht ermittelt werden.");
            }

            foreach (XElement ink in roots.SelectMany(root =>
                root.DescendantsAndSelf(One + "InkDrawing")))
            {
                XElement position = ink.Element(One + "Position");
                double x;
                double y;
                if (position == null ||
                    !TryReadCoordinate((string)position.Attribute("x"), out x) ||
                    !TryReadCoordinate((string)position.Attribute("y"), out y))
                {
                    throw new InvalidOperationException(
                        "Der gespeicherte Schnellbaustein enthält eine ungültige InkDrawing-Position.");
                }

                position.SetAttributeValue("x", FormatWidth(targetX + x));
                position.SetAttributeValue("y", FormatWidth(targetY + y));
                if (rebaseZ)
                {
                    double z;
                    if (!TryReadCoordinate((string)position.Attribute("z"), out z))
                    {
                        throw new InvalidOperationException(
                            "Der gespeicherte Schnellbaustein enthält eine ungültige InkDrawing-Position.");
                    }
                    position.SetAttributeValue("z", FormatWidth(targetZ + z));
                }
            }
        }

        private static bool TryReadCoordinate(string value, out double result)
        {
            return double.TryParse(
                    value,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out result) &&
                !double.IsNaN(result) &&
                !double.IsInfinity(result);
        }

        private static bool IsPositiveFinite(double value)
        {
            return value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static string FormatWidth(double width)
        {
            return width.ToString("0.################", CultureInfo.InvariantCulture);
        }

        private static void ExpandOutlineToMinimumWidth(
            XElement outline,
            XDocument fragment)
        {
            if (outline == null || fragment.Root == null) return;
            XElement geometry = fragment.Root.Element(Snippet + "Geometry");
            double minimumWidth;
            if (geometry == null ||
                !TryReadPositiveDouble(
                    (string)geometry.Attribute("minWidth"),
                    out minimumWidth))
            {
                return;
            }

            XElement size = outline.Element(One + "Size");
            double currentWidth;
            if (!TryReadWidth(size, out currentWidth) ||
                currentWidth >= minimumWidth)
            {
                return;
            }
            size.SetAttributeValue("width", FormatWidth(minimumWidth));
        }

        private static void ResolveBinaryData(XElement content, Func<string, string> binaryResolver)
        {
            foreach (XElement image in content.DescendantsAndSelf(One + "Image"))
            {
                ResolveBinaryObject(image, binaryResolver, "Bild");
            }
            foreach (XElement ink in content.DescendantsAndSelf(One + "InkDrawing"))
            {
                XElement file = ink.Element(One + "File");
                if (file != null) file.Remove();
                ResolveBinaryObject(ink, binaryResolver, "Freihandzeichnung");
            }
        }

        private static void ResolveBinaryObject(
            XElement element,
            Func<string, string> binaryResolver,
            string description)
        {
            XElement dataElement = element.Element(One + "Data");
            if (dataElement != null && !string.IsNullOrWhiteSpace(dataElement.Value))
            {
                XAttribute existingCallback = element.Attribute("callbackID");
                if (existingCallback != null) existingCallback.Remove();
                return;
            }
            if (dataElement != null) dataElement.Remove();

            string callback = (string)element.Attribute("callbackID");
            if (string.IsNullOrWhiteSpace(callback))
            {
                throw new InvalidOperationException(
                    description + " enthält weder eigenständige Data noch callbackID.");
            }
            if (binaryResolver == null)
            {
                throw new InvalidOperationException(
                    "Für " + description + " mit CallbackID ist kein Binärdatenleser verfügbar.");
            }
            string data = binaryResolver(callback);
            if (string.IsNullOrWhiteSpace(data))
            {
                throw new InvalidOperationException(
                    "OneNote lieferte keine Binärdaten für " + description + " " + callback + ".");
            }
            element.Add(new XElement(One + "Data", data));
            element.Attribute("callbackID").Remove();
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
