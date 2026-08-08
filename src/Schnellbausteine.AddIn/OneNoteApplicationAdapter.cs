using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;

namespace Schnellbausteine.AddIn
{
    internal sealed class OneNoteWindowInfo
    {
        public string PageId;
        public IntPtr WindowHandle;
    }

    internal interface IOneNoteApplicationAdapter : IDisposable
    {
        OneNoteWindowInfo GetCurrentWindow();
        string GetPageContent(string pageId, bool includeBinaryData, bool includeSelection);
        string GetBinaryPageContent(string pageId, string callbackId);
        void UpdatePageContent(string xml, DateTime expectedLastModifiedTime);
    }

    internal sealed class OneNoteApplicationAdapter : IOneNoteApplicationAdapter
    {
        private const XMLSchema Schema2013 = (XMLSchema)2;
        private IApplication application;

        public OneNoteApplicationAdapter()
        {
            Log.Trace("onenote.application.activate", "before");
            Type applicationType = Type.GetTypeFromProgID("OneNote.Application");
            if (applicationType == null)
            {
                throw new InvalidOperationException("OneNote Desktop ist nicht installiert.");
            }
            application = (IApplication)Activator.CreateInstance(applicationType);
            Log.Trace("onenote.application.activate", "after");
        }

        public OneNoteWindowInfo GetCurrentWindow()
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
                IntPtr windowHandle = new IntPtr(unchecked((long)window.WindowHandle));
                Log.Trace("onenote.window.WindowHandle", "after");
                return new OneNoteWindowInfo
                {
                    PageId = pageId,
                    WindowHandle = windowHandle
                };
            }
            finally
            {
                Release("onenote.window.release", window);
                Release("onenote.windows.release", windows);
            }
        }

        public string GetPageContent(
            string pageId,
            bool includeBinaryData,
            bool includeSelection)
        {
            PageInfo pageInfo = includeBinaryData && includeSelection
                ? PageInfo.piBinaryDataSelection
                : PageInfo.piBasic;
            if (includeBinaryData && !includeSelection) pageInfo = PageInfo.piBinaryData;
            if (!includeBinaryData && includeSelection) pageInfo = PageInfo.piSelection;

            string xml;
            application.GetPageContent(pageId, out xml, pageInfo, Schema2013);
            return xml;
        }

        public string GetBinaryPageContent(string pageId, string callbackId)
        {
            string value;
            application.GetBinaryPageContent(pageId, callbackId, out value);
            return value;
        }

        public void UpdatePageContent(string xml, DateTime expectedLastModifiedTime)
        {
            application.UpdatePageContent(
                xml,
                expectedLastModifiedTime,
                Schema2013,
                false);
        }

        public void Dispose()
        {
            if (application == null) return;
            Release("onenote.application.release", application);
            application = null;
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
