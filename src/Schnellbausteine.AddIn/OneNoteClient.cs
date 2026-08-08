using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    internal sealed class OneNoteClient : IDisposable
    {
        private const XMLSchema Schema2013 = (XMLSchema)2;
        private IApplication application;
        private string currentPageId;
        private IntPtr currentWindowHandle;
        private DateTime expectedLastModifiedTime;
        private bool hasExpectedLastModifiedTime;

        public OneNoteClient()
        {
            Log.Trace("onenote.client", "constructor.entry");
            Type applicationType = Type.GetTypeFromProgID("OneNote.Application");
            if (applicationType == null)
            {
                throw new InvalidOperationException("OneNote Desktop ist nicht installiert.");
            }
            Log.Trace("onenote.application.activate", "before");
            application = (IApplication)Activator.CreateInstance(applicationType);
            Log.Trace("onenote.application.activate", "after");
        }

        public XDocument ReadCurrentPage()
        {
            Log.Trace("onenote.page.read", "entry");
            currentPageId = GetCurrentPageId();
            string xml;
            Log.Trace("onenote.application.GetPageContent", "before");
            application.GetPageContent(
                currentPageId, out xml, PageInfo.piBinaryDataSelection, Schema2013);
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
                    "OneNote lieferte keinen gultigen Anderungszeitpunkt fur die aktive Seite.");
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
            string value;
            Log.Trace("onenote.application.GetBinaryPageContent", "before");
            application.GetBinaryPageContent(currentPageId, callbackId, out value);
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
                    "Die Seite wurde vor dem Aktualisieren nicht mit Anderungszeitpunkt gelesen.");
            }
            Log.Trace("onenote.application.UpdatePageContent", "before");
            application.UpdatePageContent(
                page.ToString(SaveOptions.DisableFormatting),
                expectedLastModifiedTime,
                Schema2013,
                false);
            Log.Trace("onenote.application.UpdatePageContent", "after");
            hasExpectedLastModifiedTime = false;
            Log.Trace("onenote.page.update", "exit");
        }

        private string GetCurrentPageId()
        {
            Microsoft.Office.Interop.OneNote.Windows windows = null;
            Window window = null;
            try
            {
                Log.Trace("onenote.application.Windows", "before");
                windows = application.Windows;
                Log.Trace("onenote.application.Windows", "after");
                Log.Trace("onenote.windows.CurrentWindow", "before");
                window = windows.CurrentWindow;
                Log.Trace("onenote.windows.CurrentWindow", "after");
                if (window == null)
                {
                    throw new InvalidOperationException("OneNote hat keine aktive Seite.");
                }
                Log.Trace("onenote.window.CurrentPageId", "before");
                string pageId = window.CurrentPageId;
                Log.Trace("onenote.window.CurrentPageId", "after");
                if (string.IsNullOrEmpty(pageId))
                {
                    throw new InvalidOperationException("OneNote hat keine aktive Seite.");
                }
                Log.Trace("onenote.window.WindowHandle", "before");
                currentWindowHandle = new IntPtr(unchecked((long)window.WindowHandle));
                Log.Trace("onenote.window.WindowHandle", "after");
                Log.Trace("onenote.current-page", "resolved");
                return pageId;
            }
            finally
            {
                Release("onenote.window.release", window);
                Release("onenote.windows.release", windows);
            }
        }

        public void Dispose()
        {
            Log.Trace("onenote.client", "dispose.entry");
            if (application != null)
            {
                Release("onenote.application.release", application);
                application = null;
            }
            Log.Trace("onenote.client", "dispose.exit");
        }

        private static void Release(string operation, object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            Log.Trace(operation, "before");
            Marshal.FinalReleaseComObject(value);
            Log.Trace(operation, "after");
        }
    }
}
