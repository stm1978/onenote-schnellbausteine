using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    internal sealed class OneNoteFocusException : Exception
    {
        public OneNoteFocusException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    internal sealed class OneNoteClient : IDisposable
    {
        private IOneNoteApplicationAdapter application;
        private string currentPageId;
        private IntPtr currentWindowHandle;
        private DateTime expectedLastModifiedTime;
        private bool hasExpectedLastModifiedTime;

        public OneNoteClient()
            : this(new OneNoteApplicationAdapter())
        {
            Log.Trace("onenote.client", "constructor.entry");
        }

        internal OneNoteClient(IOneNoteApplicationAdapter application)
        {
            if (application == null) throw new ArgumentNullException("application");
            this.application = application;
        }

        internal OneNoteClient(
            IOneNoteApplicationAdapter application,
            string currentPageId,
            IntPtr currentWindowHandle,
            DateTime expectedLastModifiedTime)
            : this(application)
        {
            if (string.IsNullOrEmpty(currentPageId))
            {
                throw new ArgumentException("Die Seiten-ID darf nicht leer sein.", "currentPageId");
            }
            this.currentPageId = currentPageId;
            this.currentWindowHandle = currentWindowHandle;
            this.expectedLastModifiedTime = expectedLastModifiedTime;
            hasExpectedLastModifiedTime = true;
        }

        public XDocument ReadCurrentPage()
        {
            Log.Trace("onenote.page.read", "entry");
            currentPageId = GetCurrentPageId();
            Log.Trace("onenote.application.GetPageContent", "before");
            string xml = application.GetPageContent(currentPageId, true, true);
            Log.Trace("onenote.application.GetPageContent", "after");
            XDocument page = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            XAttribute lastModified = page.Root == null
                ? null
                : page.Root.Attribute("lastModifiedTime");
            DateTime parsedLastModified;
            if (lastModified == null || !DateTime.TryParse(
                lastModified.Value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out parsedLastModified))
            {
                throw new InvalidOperationException(
                    "OneNote lieferte keinen gültigen Änderungszeitpunkt für die aktive Seite.");
            }
            expectedLastModifiedTime = parsedLastModified;
            hasExpectedLastModifiedTime = true;
            Log.Diagnostic("onenote.selection", OneNoteXml.AnalyzeSelection(page).ToDiagnosticString());
            Log.Trace("onenote.page.read", "exit");
            return page;
        }

        public IntPtr CurrentWindowHandle
        {
            get { return currentWindowHandle; }
        }

        public string ReadBinary(string callbackId)
        {
            Log.Trace("onenote.binary.read", "entry");
            if (string.IsNullOrEmpty(currentPageId)) currentPageId = GetCurrentPageId();
            Log.Trace("onenote.application.GetBinaryPageContent", "before");
            string value = application.GetBinaryPageContent(currentPageId, callbackId);
            Log.Trace("onenote.application.GetBinaryPageContent", "after");
            Log.Trace("onenote.binary.read", "exit");
            return value;
        }

        public void UpdateCurrentPage(XDocument page)
        {
            Log.Trace("onenote.page.update", "entry");
            if (string.IsNullOrEmpty(currentPageId)) currentPageId = GetCurrentPageId();
            if (!hasExpectedLastModifiedTime)
            {
                throw new InvalidOperationException(
                    "Die Seite wurde vor dem Aktualisieren nicht mit Änderungszeitpunkt gelesen.");
            }
            OneNoteXml.ValidatePostInsertCaret(page);
            Log.Trace("onenote.application.UpdatePageContent", "before");
            application.UpdatePageContent(
                page.ToString(SaveOptions.DisableFormatting),
                expectedLastModifiedTime);
            Log.Trace("onenote.application.UpdatePageContent", "after");
            hasExpectedLastModifiedTime = false;
            Log.Trace("onenote.page.update", "exit");
        }

        public void UpdateCurrentPageAndRestoreFocus(XDocument page)
        {
            UpdateCurrentPage(page);
            try
            {
                Log.Trace("onenote.application.GetPageContent.focus", "before");
                string refreshedXml = application.GetPageContent(currentPageId, false, true);
                Log.Trace("onenote.application.GetPageContent.focus", "after");
                XDocument refreshedPage = XDocument.Parse(
                    refreshedXml,
                    LoadOptions.PreserveWhitespace);
                string targetObjectId =
                    OneNoteXml.GetPostInsertFocusTargetObjectId(refreshedPage);
                Log.Diagnostic("onenote.focus.target", "resolved");
                Log.Trace("onenote.application.NavigateTo", "before");
                application.NavigateTo(currentPageId, targetObjectId);
                Log.Trace("onenote.application.NavigateTo", "after");
            }
            catch (COMException exception)
            {
                throw new OneNoteFocusException(
                    "OneNote konnte den eingefügten Cursoranker nicht fokussieren.",
                    exception);
            }
            catch (XmlException exception)
            {
                throw new OneNoteFocusException(
                    "OneNote lieferte nach dem Einfügen ungültiges Seiten-XML.",
                    exception);
            }
            catch (InvalidOperationException exception)
            {
                throw new OneNoteFocusException(
                    "OneNote lieferte nach dem Einfügen keinen eindeutigen Fokusanker.",
                    exception);
            }
        }

        private string GetCurrentPageId()
        {
            Log.Trace("onenote.current-window", "before");
            OneNoteWindowInfo window = application.GetCurrentWindow();
            Log.Trace("onenote.current-window", "after");
            if (window == null || string.IsNullOrEmpty(window.PageId))
            {
                throw new InvalidOperationException("OneNote hat keine aktive Seite.");
            }
            currentWindowHandle = window.WindowHandle;
            Log.Trace("onenote.current-page", "resolved");
            return window.PageId;
        }

        public void Dispose()
        {
            Log.Trace("onenote.client", "dispose.entry");
            if (application != null)
            {
                application.Dispose();
                application = null;
            }
            Log.Trace("onenote.client", "dispose.exit");
        }
    }
}
