using System;
using System.Windows.Forms;
using System.Xml.Linq;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    internal interface IInsertSnippetPicker : IDisposable
    {
        bool TrySelect(IWin32Window owner, out Snippet snippet);
    }

    internal static class InsertSnippetOperation
    {
        public static bool Execute(
            OneNoteClient oneNote,
            XDocument anchoredPage,
            IWin32Window owner,
            Func<IInsertSnippetPicker> createPicker,
            Func<string, XDocument> loadFragment,
            Action<OneNoteFocusException> reportFocusFailure)
        {
            if (oneNote == null) throw new ArgumentNullException("oneNote");
            if (anchoredPage == null) throw new ArgumentNullException("anchoredPage");
            if (owner == null) throw new ArgumentNullException("owner");
            if (createPicker == null) throw new ArgumentNullException("createPicker");
            if (loadFragment == null) throw new ArgumentNullException("loadFragment");
            if (reportFocusFailure == null)
            {
                throw new ArgumentNullException("reportFocusFailure");
            }

            Snippet snippet;
            using (IInsertSnippetPicker picker = createPicker())
            {
                if (picker == null)
                {
                    throw new InvalidOperationException(
                        "Der Schnellbaustein-Auswahldialog konnte nicht erstellt werden.");
                }
                if (!picker.TrySelect(owner, out snippet)) return false;
                if (snippet == null)
                {
                    throw new InvalidOperationException(
                        "Es wurde kein Schnellbaustein ausgewählt.");
                }
            }
            Log.Trace("ui.InsertDialog", "disposed");

            Log.Trace("repository.load", "before");
            XDocument fragment = loadFragment(snippet.Id);
            Log.Trace("repository.load", "after");
            Log.Trace("onenote.xml.insert", "before");
            OneNoteInsertionResult insertion =
                OneNoteXml.InsertAtSelection(anchoredPage, fragment);
            Log.Diagnostic("onenote.insert.anchor", insertion.ToDiagnosticString());
            Log.Trace("onenote.xml.insert", "after");

            try
            {
                oneNote.UpdateCurrentPageAndRestoreFocus(anchoredPage);
            }
            catch (OneNoteFocusException exception)
            {
                reportFocusFailure(exception);
            }
            return true;
        }
    }
}
