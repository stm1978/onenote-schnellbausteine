using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using Schnellbausteine.Core;

namespace Schnellbausteine.Tests
{
    internal static class Program
    {
        private static readonly List<Action> Tests = new List<Action>
        {
            RepositoryCrudAndSorting,
            RepositoryRejectsConflicts,
            AtomicFileReplacesContent,
            SelectionExtractsRichContentAndDependencies,
            SelectionExtractionPrunesPartialParagraphSiblings,
            SelectionExtractionPrunesPartialTableSiblings,
            SelectionStoresOutlineMinimumWidth,
            SelectionFallsBackToTableAndImageWidths,
            EmptyCursorIsNotASelection,
            SelectionDiagnosticsContainOnlyStructure,
            CapturedSelectionSurvivesLaterSelectionLoss,
            TextCursorUsesExactOutlineAndRun,
            StoredTextSnippetRemainsInline,
            InlineSelectionIsReplaced,
            MultipleSelectedRunsAreReplacedInPlace,
            PartialParagraphAncestorPreservesUnselectedRuns,
            MultiParagraphSelectionIsReplaced,
            WholeParagraphSelectionIsReplaced,
            BlockInsertionSplitsAndRemapsTags,
            BlockSelectionUsesExactChildBoundary,
            BlockSelectionPreservesSuffixTag,
            TableCellCursorStaysInCell,
            PartialCellAncestorPreservesCellContent,
            WholeTableCellIsReplaced,
            ListCursorPreservesListPosition,
            MultiParagraphListRangePreservesSuffixList,
            MultiParagraphRangePreservesSuffixTag,
            RichSelectionPreservesSiblings,
            QuickStylesAreCapturedAndRemapped,
            SpecialCursorUsesSafeFallback,
            PostInsertCaretIsExactForTextCursor,
            PostInsertCaretIsExactForRange,
            PostInsertCaretIsExactForBlock,
            PostInsertCaretIsExactForTable,
            PostInsertCaretIsExactForList,
            PostInsertCaretIsExactForImage,
            PostInsertCaretIsExactForMultipleParagraphs,
            NarrowTargetOutlineExpandsForText,
            WideTargetOutlineRemainsUnchanged,
            TargetWidthChangeIsOutlineScoped,
            MultiParagraphSnippetExpandsTarget,
            ListSnippetExpandsTarget,
            TableSnippetExpandsTarget,
            ImageSnippetExpandsTarget,
            MissingAndInvalidWidthsAreIgnored,
            WidthHandlingIsCultureInvariant,
            LegacySnippetDoesNotResizeTarget,
            UserInterfaceStringsUseGermanDiacritics,
            MissingSelectionIsRejectedWithoutMutation,
            AmbiguousSelectionIsRejectedWithoutMutation,
            PartialBlockWithoutDescendantIsRejected
        };

        private static int Main()
        {
            int failed = 0;
            foreach (Action test in Tests)
            {
                try
                {
                    test();
                    Console.WriteLine("PASS " + test.Method.Name);
                }
                catch (Exception exception)
                {
                    failed++;
                    Console.Error.WriteLine("FAIL " + test.Method.Name + ": " + exception);
                }
            }
            Console.WriteLine((Tests.Count - failed) + "/" + Tests.Count + " Tests erfolgreich.");
            return failed == 0 ? 0 : 1;
        }

        private static void RepositoryCrudAndSorting()
        {
            WithTemporaryDirectory(path =>
            {
                var repository = new SnippetRepository(path);
                Snippet z = repository.Save("Zulu", Fragment("<one:OE><one:T><![CDATA[Z]]></one:T></one:OE>"), false);
                Snippet a = repository.Save("Alpha", Fragment("<one:OE><one:T><![CDATA[A]]></one:T></one:OE>"), false);
                Equal("Alpha,Zulu", string.Join(",", repository.GetAll().Select(item => item.Name)));
                Equal("A", repository.LoadFragment(a.Id).Descendants(OneNoteXml.One + "T").Single().Value);
                repository.Rename(z.Id, "Beta");
                Equal("Alpha,Beta", string.Join(",", repository.GetAll().Select(item => item.Name)));
                repository.Delete(a.Id);
                Equal("Beta", repository.GetAll().Single().Name);
            });
        }

        private static void RepositoryRejectsConflicts()
        {
            WithTemporaryDirectory(path =>
            {
                var repository = new SnippetRepository(path);
                repository.Save("Name", Fragment("<one:OE><one:T><![CDATA[A]]></one:T></one:OE>"), false);
                Throws<SnippetNameConflictException>(() =>
                    repository.Save(" name ", Fragment("<one:OE><one:T><![CDATA[B]]></one:T></one:OE>"), false));
                repository.Save("NAME", Fragment("<one:OE><one:T><![CDATA[B]]></one:T></one:OE>"), true);
                Equal("B", repository.LoadFragment(repository.GetAll().Single().Id).Descendants(OneNoteXml.One + "T").Single().Value);
            });
        }

        private static void AtomicFileReplacesContent()
        {
            WithTemporaryDirectory(path =>
            {
                string file = Path.Combine(path, "value.txt");
                AtomicFile.WriteAllText(file, "alt");
                AtomicFile.WriteAllText(file, "neu");
                Equal("neu", File.ReadAllText(file));
                Equal(0, Directory.GetFiles(path, "*.tmp").Length);
            });
        }

        private static void SelectionExtractsRichContentAndDependencies()
        {
            XDocument page = Page(
                "<one:TagDef index=\"2\" type=\"0\" symbol=\"3\"/>" +
                "<one:Outline><one:OEChildren><one:OE objectID=\"old\"><one:Tag index=\"2\"/>" +
                "<one:Image callbackID=\"image1\" selected=\"all\"/></one:OE></one:OEChildren></one:Outline>");
            XDocument snippet = OneNoteXml.ExtractSelection(page, callback => callback == "image1" ? "BASE64" : null);
            Equal("BASE64", snippet.Descendants(OneNoteXml.One + "Data").Single().Value);
            Equal("2", (string)snippet.Descendants(OneNoteXml.One + "TagDef").Single().Attribute("index"));
            Assert(!snippet.Root.DescendantsAndSelf().Attributes("objectID").Any(), "objectID wurde nicht entfernt.");
        }

        private static void SelectionExtractionPrunesPartialParagraphSiblings()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T><![CDATA[vor]]></one:T><one:Image callbackID=\"unselected\"/>" +
                "<one:T selected=\"partial\"><![CDATA[auswahl]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");
            XDocument snippet = OneNoteXml.ExtractSelection(page, callback =>
            {
                throw new Exception("Unselektiertes Bild wurde aufgelost: " + callback);
            });
            XmlEqualIgnoringNamespaceDeclarations(
                Fragment("<one:OE><one:T><![CDATA[auswahl]]></one:T></one:OE>"),
                snippet);
        }

        private static void SelectionExtractionPrunesPartialTableSiblings()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table>" +
                "<one:Columns><one:Column index=\"0\" width=\"100\"/><one:Column index=\"1\" width=\"120\"/></one:Columns><one:Row>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[andere Zelle]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "<one:Cell selected=\"partial\"><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[auswahl]]></one:T>" +
                "<one:Image callbackID=\"unselected\"/></one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>");
            XDocument snippet = OneNoteXml.ExtractSelection(page, callback =>
            {
                throw new Exception("Unselektiertes Tabellenbild wurde aufgelost: " + callback);
            });
            XmlEqualIgnoringNamespaceDeclarations(
                FragmentWithMinimumWidth(
                    "<one:OE><one:Table>" +
                    "<one:Columns><one:Column index=\"0\" width=\"100\"/><one:Column index=\"1\" width=\"120\"/></one:Columns><one:Row>" +
                    "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                    "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[auswahl]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                    "</one:Row></one:Table></one:OE>",
                    "220"),
                snippet);
        }

        private static void SelectionStoresOutlineMinimumWidth()
        {
            XDocument source = Page(
                "<one:Outline><one:Position x=\"72\" y=\"96\" z=\"0\"/>" +
                "<one:Size width=\"320.5\" height=\"47.25\"/><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[Text]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");

            XDocument snippet = OneNoteXml.ExtractSelection(source, null);

            XmlEqualIgnoringNamespaceDeclarations(
                FragmentWithMinimumWidth(
                    "<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>",
                    "320.5"),
                snippet);
            Assert(!snippet.Descendants(OneNoteXml.One + "Position").Any(),
                "Die Quellposition darf nicht im Fragment landen.");
            Assert(!snippet.Descendants(OneNoteXml.One + "Size").Any(),
                "Die Outline-Geometrie darf nicht als Inhalt dupliziert werden.");
        }

        private static void SelectionFallsBackToTableAndImageWidths()
        {
            XDocument tableSource = Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table selected=\"all\">" +
                "<one:Columns><one:Column index=\"0\" width=\"100.25\"/>" +
                "<one:Column index=\"1\" width=\"120.5\"/></one:Columns>" +
                "<one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Tabelle]]></one:T>" +
                "</one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>" +
                "</one:OEChildren></one:Outline>");
            XDocument imageSource = Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:Image selected=\"all\" format=\"png\"><one:Size width=\"240.75\" height=\"80\"/>" +
                "<one:Data>BASE64</one:Data></one:Image></one:OE></one:OEChildren></one:Outline>");

            XDocument tableSnippet = OneNoteXml.ExtractSelection(tableSource, null);
            XDocument imageSnippet = OneNoteXml.ExtractSelection(imageSource, null);

            Equal("220.75", (string)tableSnippet.Root
                .Element(OneNoteXml.Snippet + "Geometry").Attribute("minWidth"));
            Equal("240.75", (string)imageSnippet.Root
                .Element(OneNoteXml.Snippet + "Geometry").Attribute("minWidth"));
        }

        private static void InlineSelectionIsReplaced()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor ]]></one:T><one:T selected=\"all\"><![CDATA[alt]]></one:T><one:T><![CDATA[ nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            XmlEqual(Page("<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor ]]></one:T><one:T><![CDATA[neu]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[ nach]]></one:T></one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void EmptyCursorIsNotASelection()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Outline>");
            Assert(!OneNoteXml.HasSelection(page), "Ein leerer Cursor wurde als Inhalt behandelt.");
            Throws<InvalidOperationException>(() => OneNoteXml.ExtractSelection(page, null));
        }

        private static void SelectionDiagnosticsContainOnlyStructure()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[geheim]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[privat]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteSelectionInfo selection = OneNoteXml.AnalyzeSelection(page);
            Equal("TextRange", selection.Scope);
            Equal("Outline", selection.ContainerScope);
            Equal(1, selection.MarkerCount);
            Equal("T:1", selection.MarkerTypes);
            Equal("/Page[1]/Outline[1]/OEChildren[1]/OE[1]/T[2]", selection.AnchorPath);
            Assert(!selection.ToDiagnosticString().Contains("geheim"), "Diagnose enthalt unselektierten Nutztext.");
            Assert(!selection.ToDiagnosticString().Contains("privat"), "Diagnose enthalt selektierten Nutztext.");
        }

        private static void CapturedSelectionSurvivesLaterSelectionLoss()
        {
            XDocument livePage = Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[erstes Outline]]></one:T></one:OE></one:OEChildren></one:Outline>" +
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            var anchoredPage = new XDocument(livePage);
            livePage.Descendants().Attributes("selected").Remove();

            OneNoteInsertionResult result = OneNoteXml.InsertAtSelection(
                anchoredPage,
                Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));

            Equal("/Page[1]/Outline[2]/OEChildren[1]/OE[1]/T[2]", result.Selection.AnchorPath);
            Equal("Inline", result.Strategy);
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[erstes Outline]]></one:T></one:OE></one:OEChildren></one:Outline>" +
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T><one:T><![CDATA[neu]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>"),
                anchoredPage);
            Throws<InvalidOperationException>(() => OneNoteXml.InsertAtSelection(
                livePage,
                Fragment("<one:OE><one:T><![CDATA[falsch]]></one:T></one:OE>")));
        }

        private static void TextCursorUsesExactOutlineAndRun()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[nicht hier]]></one:T></one:OE></one:OEChildren></one:Outline>" +
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[links]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[rechts]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.InsertAtSelection(page, Fragment("<one:OE><one:T><![CDATA[Mitte]]></one:T></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[nicht hier]]></one:T></one:OE></one:OEChildren></one:Outline>" +
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[links]]></one:T><one:T><![CDATA[Mitte]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[rechts]]></one:T></one:OE></one:OEChildren></one:Outline>"),
                page);
        }

        private static void StoredTextSnippetRemainsInline()
        {
            WithTemporaryDirectory(path =>
            {
                XDocument source = Page(
                    "<one:Outline><one:OEChildren><one:OE>" +
                    "<one:T selected=\"all\"><![CDATA[gespeichert]]></one:T>" +
                    "</one:OE></one:OEChildren></one:Outline>");
                XDocument extracted = OneNoteXml.ExtractSelection(source, null);
                var repository = new SnippetRepository(path);
                Snippet snippet = repository.Save("Text", extracted, false);
                XDocument loaded = repository.LoadFragment(snippet.Id);
                Assert(
                    loaded.Descendants(OneNoteXml.One + "OE").Single()
                        .Attributes().Any(attribute => attribute.IsNamespaceDeclaration),
                    "Der Roundtrip enthalt keine explizite Namespace-Deklaration.");

                XDocument target = Page(
                    "<one:Outline><one:OEChildren><one:OE>" +
                    "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                    "</one:OE></one:OEChildren></one:Outline>");
                OneNoteInsertionResult result = OneNoteXml.InsertAtSelection(target, loaded);

                Equal("Inline", result.Strategy);
                XmlEqual(Page(
                    "<one:Outline><one:OEChildren><one:OE>" +
                    "<one:T><![CDATA[vor]]></one:T><one:T><![CDATA[gespeichert]]></one:T>" +
                    "<one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                    "</one:OE></one:OEChildren></one:Outline>"), target);
            });
        }

        private static void MultipleSelectedRunsAreReplacedInPlace()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T>" +
                "<one:T selected=\"partial\"><![CDATA[a]]></one:T><one:T selected=\"all\"><![CDATA[b]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.InsertAtSelection(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T>" +
                "<one:T><![CDATA[neu]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>"),
                page);
        }

        private static void PartialParagraphAncestorPreservesUnselectedRuns()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"partial\"><![CDATA[alt]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteInsertionResult result = OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("/Page[1]/Outline[1]/OEChildren[1]/OE[1]/T[2]", result.Selection.AnchorPath);
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[vor]]></one:T><one:T><![CDATA[neu]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void BlockInsertionSplitsAndRemapsTags()
        {
            XDocument page = Page(
                "<one:TagDef index=\"0\" type=\"0\" symbol=\"1\"/>" +
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[alt]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            XDocument fragment = Fragment("<one:OE><one:Tag index=\"7\"/><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Tabelle]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>",
                "<one:TagDef index=\"7\" type=\"0\" symbol=\"3\"/>");
            OneNoteXml.Insert(page, fragment);
            Equal("vorTabelle" + string.Empty + "nach", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
            Equal("1", (string)page.Descendants(OneNoteXml.One + "Tag").Single().Attribute("index"));
            Equal(2, page.Root.Elements(OneNoteXml.One + "TagDef").Count());
            Equal("Table", page.Descendants(OneNoteXml.One + "OEChildren").First().Elements(OneNoteXml.One + "OE").ElementAt(1).Elements().Last().Name.LocalName);
        }

        private static void WholeParagraphSelectionIsReplaced()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE selected=\"all\"><one:T><![CDATA[alt]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            XmlEqual(Page("<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[neu]]></one:T></one:OE><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void MultiParagraphSelectionIsReplaced()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[erst]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[zweit]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T></one:OE>" +
                "<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void BlockSelectionUsesExactChildBoundary()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[vor]]></one:T><one:Image selected=\"all\"><one:Data>OLD</one:Data></one:Image>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[neu]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T></one:OE>" +
                "<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[neu]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void BlockSelectionPreservesSuffixTag()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:Tag index=\"3\"/>" +
                "<one:T selected=\"all\"><![CDATA[alt]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[block]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[block]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>" +
                "<one:OE><one:Tag index=\"3\"/><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void TableCellCursorStaysInCell()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[A]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>");
            OneNoteInsertionResult result = OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("TableCell", result.Selection.ContainerScope);
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[A]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T><one:T><![CDATA[neu]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void PartialCellAncestorPreservesCellContent()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell selected=\"partial\"><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>");
            OneNoteInsertionResult result = OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("TableCell", result.Selection.ContainerScope);
            Equal("Inline", result.Strategy);
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[vor]]></one:T><one:T><![CDATA[neu]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void WholeTableCellIsReplaced()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell selected=\"all\"><one:OEChildren><one:OE><one:T><![CDATA[alt]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[bleibt]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>");
            OneNoteInsertionResult result = OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("ReplaceTableCell", result.Strategy);
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[neu]]></one:T></one:OE><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[bleibt]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void ListCursorPreservesListPosition()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:List><one:Bullet bullet=\"2\"/></one:List>" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.InsertAtSelection(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:List><one:Bullet bullet=\"2\"/></one:List>" +
                "<one:T><![CDATA[vor]]></one:T><one:T><![CDATA[neu]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void MultiParagraphListRangePreservesSuffixList()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:List><one:Bullet bullet=\"2\"/></one:List><one:T><![CDATA[vor]]></one:T><one:T selected=\"partial\"><![CDATA[ab]]></one:T></one:OE>" +
                "<one:OE><one:List><one:Bullet bullet=\"2\"/></one:List><one:T selected=\"partial\"><![CDATA[bis]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[block]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:List><one:Bullet bullet=\"2\"/></one:List><one:T><![CDATA[vor]]></one:T></one:OE>" +
                "<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[block]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>" +
                "<one:OE><one:List><one:Bullet bullet=\"2\"/></one:List><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void MultiParagraphRangePreservesSuffixTag()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T><one:T selected=\"partial\"><![CDATA[ab]]></one:T></one:OE>" +
                "<one:OE><one:Tag index=\"4\"/><one:T selected=\"partial\"><![CDATA[bis]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[block]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T></one:OE>" +
                "<one:OE><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[block]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>" +
                "<one:OE><one:Tag index=\"4\"/><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void SpecialCursorUsesSafeFallback()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[<!--MathML--><a href=\"x\"></a>]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteInsertionResult result = OneNoteXml.InsertAtSelection(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("AfterSpecialParagraph", result.Strategy);
            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[<!--MathML--><a href=\"x\"></a>]]></one:T></one:OE>" +
                "<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void PostInsertCaretIsExactForTextCursor()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[vor]]></one:T><one:T><![CDATA[Text]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
            AssertSinglePostInsertCaret(page);
        }

        private static void PostInsertCaretIsExactForRange()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"partial\"><![CDATA[alt]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[vor]]></one:T><one:T><![CDATA[Text]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[nach]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
            AssertSinglePostInsertCaret(page);
        }

        private static void PostInsertCaretIsExactForBlock()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[alt]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:Table><one:Row><one:Cell><one:OEChildren>" +
                    "<one:OE><one:T><![CDATA[Tabelle]]></one:T></one:OE>" +
                    "</one:OEChildren></one:Cell></one:Row></one:Table></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T></one:OE>" +
                "<one:OE><one:Table><one:Row><one:Cell><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[Tabelle]]></one:T></one:OE>" +
                "</one:OEChildren></one:Cell></one:Row></one:Table></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
            AssertSinglePostInsertCaret(page);
        }

        private static void PostInsertCaretIsExactForTable()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell selected=\"partial\"><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:Table><one:Row><one:Cell><one:OEChildren>" +
                    "<one:OE><one:T><![CDATA[innen]]></one:T></one:OE>" +
                    "</one:OEChildren></one:Cell></one:Row></one:Table></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:Table><one:Row>" +
                "<one:Cell><one:OEChildren>" +
                "<one:OE><one:Table><one:Row><one:Cell><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[innen]]></one:T></one:OE>" +
                "</one:OEChildren></one:Cell></one:Row></one:Table></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE></one:OEChildren></one:Outline>"), page);
            AssertSinglePostInsertCaret(page);
        }

        private static void PostInsertCaretIsExactForList()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:List><one:Bullet bullet=\"2\"/></one:List>" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[Listentext]]></one:T></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:List><one:Bullet bullet=\"2\"/></one:List>" +
                "<one:T><![CDATA[vor]]></one:T><one:T><![CDATA[Listentext]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
            AssertSinglePostInsertCaret(page);
        }

        private static void PostInsertCaretIsExactForImage()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:Image format=\"png\"><one:Data>BASE64</one:Data></one:Image></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T></one:OE>" +
                "<one:OE><one:Image format=\"png\"><one:Data>BASE64</one:Data></one:Image></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "<one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
            AssertSinglePostInsertCaret(page);
        }

        private static void PostInsertCaretIsExactForMultipleParagraphs()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment(
                    "<one:OE><one:T><![CDATA[Absatz 1]]></one:T></one:OE>" +
                    "<one:OE><one:T><![CDATA[Absatz 2]]></one:T></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[Absatz 1]]></one:T></one:OE>" +
                "<one:OE><one:T><![CDATA[Absatz 2]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
            AssertSinglePostInsertCaret(page);
        }

        private static void NarrowTargetOutlineExpandsForText()
        {
            XDocument page = Page(
                "<one:Outline><one:Position x=\"10\" y=\"20\" z=\"0\"/>" +
                "<one:Size width=\"80\" height=\"18\"/><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                FragmentWithMinimumWidth(
                    "<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>",
                    "160.5"));

            XmlEqual(Page(
                "<one:Outline><one:Position x=\"10\" y=\"20\" z=\"0\"/>" +
                "<one:Size width=\"160.5\" height=\"18\"/><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Text]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void WideTargetOutlineRemainsUnchanged()
        {
            XDocument page = Page(
                "<one:Outline><one:Size width=\"300.25\" height=\"18\"/>" +
                "<one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                FragmentWithMinimumWidth(
                    "<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>",
                    "160.5"));

            XmlEqual(Page(
                "<one:Outline><one:Size width=\"300.25\" height=\"18\"/>" +
                "<one:OEChildren><one:OE><one:T><![CDATA[Text]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void TargetWidthChangeIsOutlineScoped()
        {
            XDocument page = Page(
                "<one:Outline><one:Size width=\"40\" height=\"10\"/><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Andere Outline]]></one:T></one:OE></one:OEChildren></one:Outline>" +
                "<one:Outline><one:Size width=\"60\" height=\"20\"/><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                FragmentWithMinimumWidth(
                    "<one:OE><one:T><![CDATA[Ziel]]></one:T></one:OE>",
                    "180"));

            XmlEqual(Page(
                "<one:Outline><one:Size width=\"40\" height=\"10\"/><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Andere Outline]]></one:T></one:OE></one:OEChildren></one:Outline>" +
                "<one:Outline><one:Size width=\"180\" height=\"20\"/><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Ziel]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void MultiParagraphSnippetExpandsTarget()
        {
            XDocument page = Page(
                "<one:Outline><one:Size width=\"90\" height=\"40\"/><one:OEChildren>" +
                "<one:OE selected=\"partial\"><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                FragmentWithMinimumWidth(
                    "<one:OE><one:T><![CDATA[Absatz 1]]></one:T></one:OE>" +
                    "<one:OE><one:T><![CDATA[Absatz 2]]></one:T></one:OE>",
                    "260"));

            XmlEqual(Page(
                "<one:Outline><one:Size width=\"260\" height=\"40\"/><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[Absatz 1]]></one:T></one:OE>" +
                "<one:OE><one:T><![CDATA[Absatz 2]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void ListSnippetExpandsTarget()
        {
            XDocument page = Page(
                "<one:Outline><one:Size width=\"100\" height=\"24\"/><one:OEChildren>" +
                "<one:OE selected=\"partial\"><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                FragmentWithMinimumWidth(
                    "<one:OE><one:List><one:Bullet bullet=\"2\"/></one:List>" +
                    "<one:T><![CDATA[Listenpunkt]]></one:T></one:OE>",
                    "210"));

            XmlEqual(Page(
                "<one:Outline><one:Size width=\"210\" height=\"24\"/><one:OEChildren>" +
                "<one:OE><one:List><one:Bullet bullet=\"2\"/></one:List>" +
                "<one:T><![CDATA[Listenpunkt]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void TableSnippetExpandsTarget()
        {
            XDocument page = Page(
                "<one:Outline><one:Size width=\"120\" height=\"50\"/><one:OEChildren>" +
                "<one:OE selected=\"partial\"><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            string table =
                "<one:OE><one:Table><one:Columns>" +
                "<one:Column index=\"0\" width=\"130\"/><one:Column index=\"1\" width=\"140\"/>" +
                "</one:Columns><one:Row><one:Cell><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Tabelle]]></one:T></one:OE></one:OEChildren></one:Cell>" +
                "</one:Row></one:Table></one:OE>";

            OneNoteXml.InsertAtSelection(
                page,
                FragmentWithMinimumWidth(table, "270"));

            XmlEqual(Page(
                "<one:Outline><one:Size width=\"270\" height=\"50\"/><one:OEChildren>" +
                table +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void ImageSnippetExpandsTarget()
        {
            XDocument page = Page(
                "<one:Outline><one:Size width=\"110\" height=\"35\"/><one:OEChildren>" +
                "<one:OE selected=\"partial\"><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            string image =
                "<one:OE><one:Image format=\"png\"><one:Size width=\"240\" height=\"80\"/>" +
                "<one:Data>BASE64</one:Data></one:Image></one:OE>";

            OneNoteXml.InsertAtSelection(
                page,
                FragmentWithMinimumWidth(image, "240"));

            XmlEqual(Page(
                "<one:Outline><one:Size width=\"240\" height=\"35\"/><one:OEChildren>" +
                image +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void MissingAndInvalidWidthsAreIgnored()
        {
            XDocument missingTargetSize = Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            XDocument invalidTargetWidth = Page(
                "<one:Outline><one:Size width=\"unbekannt\" height=\"20\"/><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            XDocument invalidMetadataTarget = Page(
                "<one:Outline><one:Size width=\"75\" height=\"20\"/><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            XDocument invalidMetadata = FragmentWithMinimumWidth(
                "<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>",
                "keine-zahl");
            XDocument invalidSourceSize = Page(
                "<one:Outline><one:Size width=\"-12\" height=\"20\"/><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[Quelle]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                missingTargetSize,
                FragmentWithMinimumWidth(
                    "<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>",
                    "150"));
            OneNoteXml.InsertAtSelection(
                invalidTargetWidth,
                FragmentWithMinimumWidth(
                    "<one:OE><one:T><![CDATA[Text]]></one:T></one:OE>",
                    "150"));
            OneNoteXml.InsertAtSelection(invalidMetadataTarget, invalidMetadata);
            XDocument invalidSourceSnippet =
                OneNoteXml.ExtractSelection(invalidSourceSize, null);

            XmlEqual(Page(
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[Text]]></one:T>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), missingTargetSize);
            XmlEqual(Page(
                "<one:Outline><one:Size width=\"unbekannt\" height=\"20\"/><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Text]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), invalidTargetWidth);
            XmlEqual(Page(
                "<one:Outline><one:Size width=\"75\" height=\"20\"/><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Text]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), invalidMetadataTarget);
            Assert(invalidSourceSnippet.Root.Element(OneNoteXml.Snippet + "Geometry") == null,
                "Ungültige Quellbreite darf kein Geometriemetadatum erzeugen.");
        }

        private static void WidthHandlingIsCultureInvariant()
        {
            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
            CultureInfo originalUiCulture = Thread.CurrentThread.CurrentUICulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("de-DE");
                XDocument source = Page(
                    "<one:Outline><one:Size width=\"123.5\" height=\"20\"/>" +
                    "<one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[Text]]></one:T>" +
                    "</one:OE></one:OEChildren></one:Outline>");
                XDocument target = Page(
                    "<one:Outline><one:Size width=\"80.25\" height=\"20\"/>" +
                    "<one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                    "</one:OE></one:OEChildren></one:Outline>");

                XDocument snippet = OneNoteXml.ExtractSelection(source, null);
                OneNoteXml.InsertAtSelection(target, snippet);

                Equal("123.5", (string)snippet.Root
                    .Element(OneNoteXml.Snippet + "Geometry").Attribute("minWidth"));
                XmlEqual(Page(
                    "<one:Outline><one:Size width=\"123.5\" height=\"20\"/>" +
                    "<one:OEChildren><one:OE><one:T><![CDATA[Text]]></one:T>" +
                    "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                    "</one:OE></one:OEChildren></one:Outline>"), target);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = originalCulture;
                Thread.CurrentThread.CurrentUICulture = originalUiCulture;
            }
        }

        private static void LegacySnippetDoesNotResizeTarget()
        {
            XDocument page = Page(
                "<one:Outline><one:Size width=\"70\" height=\"18\"/><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");

            OneNoteXml.InsertAtSelection(
                page,
                Fragment("<one:OE><one:T><![CDATA[Altbestand]]></one:T></one:OE>"));

            XmlEqual(Page(
                "<one:Outline><one:Size width=\"70\" height=\"18\"/><one:OEChildren><one:OE>" +
                "<one:T><![CDATA[Altbestand]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>"), page);
        }

        private static void UserInterfaceStringsUseGermanDiacritics()
        {
            string root = FindRepositoryRoot();
            string[] sourceFiles = Directory.GetFiles(
                    Path.Combine(root, "src"),
                    "*.*",
                    SearchOption.AllDirectories)
                .Concat(Directory.GetFiles(
                    Path.Combine(root, "installer"),
                    "*.*",
                    SearchOption.AllDirectories))
                .Where(path =>
                    path.EndsWith(".cs", StringComparison.OrdinalIgnoreCase) ||
                    path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                    path.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase) ||
                    path.EndsWith(".iss", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            var asciiReplacement = new Regex(
                "\\b(ausfuhren|lauft|haengenden|Schliessen|schliessen|" +
                "Loschen|loschen|Einfugen|einfugen|moglich|gultig|" +
                "Anderungszeitpunkt|ausgewahlt|bestatigen|hochstens|" +
                "enthalt|verfugbar|unterstutzt|prazise|unvollstandig|" +
                "Binardaten|fur|Fur|uber|Uber|Uberschreiben|uberschreiben)\\b",
                RegexOptions.CultureInvariant);

            foreach (string path in sourceFiles)
            {
                string content = File.ReadAllText(path);
                Match match = asciiReplacement.Match(content);
                Assert(!match.Success,
                    "ASCII-Ersatzschreibweise in " + path + ": " + match.Value);
                if (path.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase) &&
                    content.Any(character => character > 127))
                {
                    byte[] bytes = File.ReadAllBytes(path);
                    Assert(bytes.Length >= 3 &&
                        bytes[0] == 0xEF &&
                        bytes[1] == 0xBB &&
                        bytes[2] == 0xBF,
                        "PowerShell-Skript mit Unicode-Text benötigt eine UTF-8-BOM: " + path);
                }
            }

            string dialogs = File.ReadAllText(
                Path.Combine(root, "src", "Schnellbausteine.AddIn", "Dialogs.cs"));
            Assert(
                dialogs.Contains("\"Schnellbausteine verwalten\""),
                "UI-Text Schnellbausteine verwalten fehlt.");
            Assert(dialogs.Contains("\"Löschen\""), "UI-Text Löschen fehlt.");
            Assert(dialogs.Contains("\"Schließen\""), "UI-Text Schließen fehlt.");
        }

        private static void QuickStylesAreCapturedAndRemapped()
        {
            XDocument source = Page(
                "<one:QuickStyleDef index=\"5\" name=\"Heading 2\" fontSize=\"16.0\"/>" +
                "<one:Outline><one:OEChildren><one:OE quickStyleIndex=\"5\"><one:T selected=\"all\"><![CDATA[Titel]]></one:T></one:OE></one:OEChildren></one:Outline>");
            XDocument fragment = OneNoteXml.ExtractSelection(source, null);
            Equal("5", (string)fragment.Descendants(OneNoteXml.One + "QuickStyleDef").Single().Attribute("index"));

            XDocument target = Page(
                "<one:QuickStyleDef index=\"0\" name=\"Normal\" fontSize=\"11.0\"/>" +
                "<one:Outline><one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(target, fragment);
            Equal("1", (string)target.Descendants(OneNoteXml.One + "OE").Single(element => (string)element.Attribute("quickStyleIndex") != null).Attribute("quickStyleIndex"));
            Equal(2, target.Root.Elements(OneNoteXml.One + "QuickStyleDef").Count());
        }

        private static void RichSelectionPreservesSiblings()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:Image callbackID=\"image1\" selected=\"all\"/>" +
                "<one:T><![CDATA[behalten]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            XmlEqual(Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[behalten]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>"), page);
        }

        private static void MissingSelectionIsRejectedWithoutMutation()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[alt]]></one:T></one:OE></one:OEChildren></one:Outline>");
            XDocument before = new XDocument(page);
            OneNoteSelectionInfo selection = OneNoteXml.AnalyzeSelection(page);
            Equal("Missing", selection.Scope);
            Assert(!selection.IsUsable, "Fehlende Selection wurde als verwendbar gemeldet.");
            Throws<InvalidOperationException>(() =>
                OneNoteXml.InsertAtSelection(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>")));
            XmlEqual(before, page);
        }

        private static void AmbiguousSelectionIsRejectedWithoutMutation()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Outline>" +
                "<one:Outline><one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Outline>");
            XDocument before = new XDocument(page);
            OneNoteSelectionInfo selection = OneNoteXml.AnalyzeSelection(page);
            Equal("Ambiguous", selection.Scope);
            Assert(selection.MarkerPaths.Contains("/Outline[1]/") &&
                selection.MarkerPaths.Contains("/Outline[2]/"), "Diagnose enthalt nicht beide Markerpfade.");
            Throws<InvalidOperationException>(() =>
                OneNoteXml.InsertAtSelection(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>")));
            XmlEqual(before, page);
        }

        private static void PartialBlockWithoutDescendantIsRejected()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren><one:OE selected=\"partial\">" +
                "<one:T><![CDATA[inhalt]]></one:T><one:Image callbackID=\"image1\"/>" +
                "</one:OE></one:OEChildren></one:Outline>");
            XDocument before = new XDocument(page);
            OneNoteSelectionInfo selection = OneNoteXml.AnalyzeSelection(page);
            Equal("Ambiguous", selection.Scope);
            Throws<InvalidOperationException>(() => OneNoteXml.ExtractSelection(page, callback => "DATA"));
            Throws<InvalidOperationException>(() =>
                OneNoteXml.InsertAtSelection(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>")));
            XmlEqual(before, page);
        }

        private static XDocument Page(string content)
        {
            return XDocument.Parse("<one:Page xmlns:one=\"" + OneNoteXml.One + "\">" + content + "</one:Page>", LoadOptions.PreserveWhitespace);
        }

        private static XDocument Fragment(string content, string definitions = "")
        {
            return XDocument.Parse("<sq:Snippet xmlns:sq=\"" + OneNoteXml.Snippet + "\" xmlns:one=\"" + OneNoteXml.One + "\" version=\"1\"><sq:TagDefs>" +
                definitions + "</sq:TagDefs><sq:QuickStyleDefs/><sq:Content>" + content + "</sq:Content></sq:Snippet>", LoadOptions.PreserveWhitespace);
        }

        private static XDocument FragmentWithMinimumWidth(
            string content,
            string minimumWidth)
        {
            XDocument fragment = Fragment(content);
            fragment.Root.AddFirst(
                new XElement(
                    OneNoteXml.Snippet + "Geometry",
                    new XAttribute("minWidth", minimumWidth)));
            return fragment;
        }

        private static void AssertSinglePostInsertCaret(XDocument page)
        {
            OneNoteXml.ValidatePostInsertCaret(page);
            XElement caret = page.Descendants(OneNoteXml.One + "T")
                .Single(element => (string)element.Attribute("selected") == "all");
            Equal(string.Empty, caret.Value);
            Equal("OE", caret.Parent.Name.LocalName);
            Equal(1, page.Root.DescendantsAndSelf().Attributes("selected").Count());
            Assert(page.ToString(SaveOptions.DisableFormatting)
                .Contains("<![CDATA[]]>"), "Der serialisierte Cursoranker enthält kein leeres CDATA.");
        }

        private static string FindRepositoryRoot()
        {
            foreach (string start in new[]
            {
                Environment.CurrentDirectory,
                AppDomain.CurrentDomain.BaseDirectory
            })
            {
                DirectoryInfo directory = new DirectoryInfo(start);
                while (directory != null)
                {
                    if (File.Exists(Path.Combine(directory.FullName, "Schnellbausteine.sln")))
                    {
                        return directory.FullName;
                    }
                    directory = directory.Parent;
                }
            }
            throw new DirectoryNotFoundException("Repository-Wurzel wurde nicht gefunden.");
        }

        private static void XmlEqual(XDocument expected, XDocument actual)
        {
            if (!XNode.DeepEquals(expected, actual))
            {
                throw new Exception(
                    "XML weicht ab.\nErwartet: " + expected.ToString(SaveOptions.DisableFormatting) +
                    "\nTatsachlich: " + actual.ToString(SaveOptions.DisableFormatting));
            }
        }

        private static void XmlEqualIgnoringNamespaceDeclarations(
            XDocument expected,
            XDocument actual)
        {
            var expectedClone = new XDocument(expected);
            var actualClone = new XDocument(actual);
            expectedClone.Root.DescendantsAndSelf().Attributes()
                .Where(attribute => attribute.IsNamespaceDeclaration)
                .Remove();
            actualClone.Root.DescendantsAndSelf().Attributes()
                .Where(attribute => attribute.IsNamespaceDeclaration)
                .Remove();
            NormalizeEmptyElements(expectedClone);
            NormalizeEmptyElements(actualClone);
            XmlEqual(expectedClone, actualClone);
        }

        private static void NormalizeEmptyElements(XDocument document)
        {
            foreach (XElement element in document.Descendants()
                .Where(candidate => !candidate.Nodes().Any())
                .Reverse()
                .ToList())
            {
                element.ReplaceWith(new XElement(element.Name, element.Attributes()));
            }
        }

        private static void WithTemporaryDirectory(Action<string> action)
        {
            string path = Path.Combine(Path.GetTempPath(), "Schnellbausteine.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            try { action(path); }
            finally { Directory.Delete(path, true); }
        }

        private static void Equal<T>(T expected, T actual)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                throw new Exception("Erwartet: " + expected + "; tatsachlich: " + actual);
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }

        private static void Throws<T>(Action action) where T : Exception
        {
            try { action(); }
            catch (T) { return; }
            throw new Exception("Ausnahme " + typeof(T).Name + " wurde erwartet.");
        }
    }
}
