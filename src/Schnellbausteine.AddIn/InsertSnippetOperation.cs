using System;
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

            Log.Trace("repository.load", "before");
            XDocument fragment = loadFragment(snippet.Id);
            Log.Trace("repository.load", "after");
            Log.Trace("onenote.xml.insert", "before");
            OneNoteInsertionResult insertion =
                OneNoteXml.InsertAtSelection(anchoredPage, fragment);
            Log.Diagnostic("onenote.insert.anchor", insertion.ToDiagnosticString());
            Log.Trace("onenote.xml.insert", "after");
            oneNote.UpdateCurrentPageAndActivateWindow(anchoredPage);
        }
    }
}
