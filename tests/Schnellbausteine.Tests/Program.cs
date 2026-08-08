using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            EmptyCursorIsNotASelection,
            InlineSelectionIsReplaced,
            MultiParagraphSelectionIsReplaced,
            WholeParagraphSelectionIsReplaced,
            BlockInsertionSplitsAndRemapsTags,
            RichSelectionPreservesSiblings,
            QuickStylesAreCapturedAndRemapped,
            SpecialCursorUsesSafeFallback,
            MissingSelectionUsesOutlineEnd
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

        private static void InlineSelectionIsReplaced()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor ]]></one:T><one:T selected=\"all\"><![CDATA[alt]]></one:T><one:T><![CDATA[ nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("vor neu nach", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
            Equal("", page.Descendants(OneNoteXml.One + "T").Single(item => (string)item.Attribute("selected") == "all").Value);
        }

        private static void EmptyCursorIsNotASelection()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[]]></one:T></one:OE></one:OEChildren></one:Outline>");
            Assert(!OneNoteXml.HasSelection(page), "Ein leerer Cursor wurde als Inhalt behandelt.");
            Throws<InvalidOperationException>(() => OneNoteXml.ExtractSelection(page, null));
        }

        private static void BlockInsertionSplitsAndRemapsTags()
        {
            XDocument page = Page(
                "<one:TagDef index=\"0\" type=\"0\" symbol=\"1\"/>" +
                "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[alt]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE></one:OEChildren></one:Outline>");
            XDocument fragment = Fragment("<one:OE><one:Tag index=\"7\"/><one:Table><one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Tabelle]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>",
                "<one:TagDef index=\"7\" type=\"0\" symbol=\"3\"/>");
            OneNoteXml.Insert(page, fragment);
            Equal("vorTabellenach", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
            Equal("1", (string)page.Descendants(OneNoteXml.One + "Tag").Single().Attribute("index"));
            Equal(2, page.Root.Elements(OneNoteXml.One + "TagDef").Count());
        }

        private static void WholeParagraphSelectionIsReplaced()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE selected=\"all\"><one:T><![CDATA[alt]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("neu", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
        }

        private static void MultiParagraphSelectionIsReplaced()
        {
            XDocument page = Page(
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[vor]]></one:T><one:T selected=\"all\"><![CDATA[erst]]></one:T></one:OE>" +
                "<one:OE><one:T selected=\"all\"><![CDATA[zweit]]></one:T><one:T><![CDATA[nach]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("vorneunach", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
        }

        private static void SpecialCursorUsesSafeFallback()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T selected=\"all\"><![CDATA[<!--MathML--><a href=\"x\"></a>]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("<!--MathML--><a href=\"x\"></a>neu", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
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
            Equal("behaltenneu", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
            Assert(!page.Descendants(OneNoteXml.One + "Image").Any(), "Das ausgewahlte Bild wurde nicht ersetzt.");
        }

        private static void MissingSelectionUsesOutlineEnd()
        {
            XDocument page = Page("<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[alt]]></one:T></one:OE></one:OEChildren></one:Outline>");
            OneNoteXml.Insert(page, Fragment("<one:OE><one:T><![CDATA[neu]]></one:T></one:OE>"));
            Equal("altneu", string.Concat(page.Descendants(OneNoteXml.One + "T").Select(item => item.Value)));
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
