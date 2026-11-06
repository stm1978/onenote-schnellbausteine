using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    internal sealed class OneNoteClient : IDisposable
    {
        private IOneNoteApplicationAdapter application;
        private readonly IOneNoteWindowActivator windowActivator;
        private string currentPageId;
        private IntPtr currentWindowHandle;
        private DateTime expectedLastModifiedTime;
        private bool hasExpectedLastModifiedTime;

        public OneNoteClient()
            : this(new OneNoteApplicationAdapter(), new OneNoteWindowActivator())
        {
            Log.Trace("onenote.client", "constructor.entry");
        }

        internal OneNoteClient(IOneNoteApplicationAdapter application)
            : this(application, new OneNoteWindowActivator())
        {
        }

        internal OneNoteClient(
            IOneNoteApplicationAdapter application,
            IOneNoteWindowActivator windowActivator)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (windowActivator == null) throw new ArgumentNullException("windowActivator");
            this.application = application;
            this.windowActivator = windowActivator;
        }

        internal OneNoteClient(
            IOneNoteApplicationAdapter application,
            string currentPageId,
            IntPtr currentWindowHandle,
            DateTime expectedLastModifiedTime,
            IOneNoteWindowActivator windowActivator)
            : this(application, windowActivator)
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
            return ReadCurrentPage(true, "insert");
        }

        public XDocument ReadCurrentPageForCapture()
        {
            return ReadCurrentPage(false, "save");
        }

        private XDocument ReadCurrentPage(bool includeBinaryData, string operation)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Log.Trace("onenote.page.read", "entry");
            currentPageId = GetCurrentPageId();
            Log.Trace("onenote.application.GetPageContent", "before");
            string xml = application.GetPageContent(currentPageId, includeBinaryData, true);
            Log.Trace("onenote.application.GetPageContent", "after");
            Log.Diagnostic(
                "performance." + operation + ".get-page-content",
                "elapsedMs=" + stopwatch.ElapsedMilliseconds);
            stopwatch.Restart();
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
            Log.Diagnostic(
                "performance." + operation + ".parse-and-analyze",
                "elapsedMs=" + stopwatch.ElapsedMilliseconds);
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
            ValidateInkFreeUpdate(page);
            Log.Trace("onenote.page.update", "entry");
            if (string.IsNullOrEmpty(currentPageId)) currentPageId = GetCurrentPageId();
            if (!hasExpectedLastModifiedTime)
            {
                throw new InvalidOperationException(
                    "Die Seite wurde vor dem Aktualisieren nicht mit Änderungszeitpunkt gelesen.");
            }
            OneNoteXml.ValidatePostInsertCaret(page);
            Stopwatch stopwatch = Stopwatch.StartNew();
            string xml = page.ToString(SaveOptions.DisableFormatting);
            Log.Diagnostic(
                "performance.insert.serialize-page",
                "elapsedMs=" + stopwatch.ElapsedMilliseconds);
            stopwatch.Restart();
            Log.Trace("onenote.application.UpdatePageContent", "before");
            application.UpdatePageContent(
                xml,
                expectedLastModifiedTime);
            Log.Trace("onenote.application.UpdatePageContent", "after");
            Log.Diagnostic(
                "performance.insert.update-page-content",
                "elapsedMs=" + stopwatch.ElapsedMilliseconds);
            hasExpectedLastModifiedTime = false;
            Log.Trace("onenote.page.update", "exit");
        }

        internal static void ValidateInkFreeUpdate(XDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            // Rewriting exported ink can discard effects on existing originals as well.
            if (document.Descendants(OneNoteXml.One + "InkDrawing").Any())
            {
                throw new InvalidOperationException(
                    "Einfügen ist für Freihand-Bausteine und auf Seiten mit Freihandzeichnungen " +
                    "vorübergehend gesperrt. Der XML-Schreibpfad kann Farbeffekte vorhandener " +
                    "Originalzeichnungen verlieren. Die Seite wurde nicht aktualisiert.");
            }
        }

        public void UpdateCurrentPageAndActivateWindow(XDocument page)
        {
            UpdateCurrentPage(page);
            try
            {
                Log.Trace("onenote.window.activate", "before");
                windowActivator.Activate(currentWindowHandle);
                Log.Trace("onenote.window.activate", "after");
            }
            catch (Exception exception)
            {
                Log.Error(
                    "Der Schnellbaustein wurde eingefügt, aber das OneNote-Hauptfenster " +
                    "konnte nicht aktiviert werden.",
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
