using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    internal sealed class OneNoteClient : IDisposable
    {
        private const XMLSchema Schema2013 = (XMLSchema)2;
        private dynamic application;
        private string currentPageId;

        public OneNoteClient()
        {
            Type applicationType = Type.GetTypeFromProgID("OneNote.Application");
            if (applicationType == null)
            {
                throw new InvalidOperationException("OneNote Desktop ist nicht installiert.");
            }
            application = Activator.CreateInstance(applicationType);
        }

        public XDocument ReadCurrentPage()
        {
            currentPageId = GetCurrentPageId();
            string xml;
            application.GetPageContent(currentPageId, out xml, PageInfo.piBinaryData, Schema2013);
            return XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        }

        public string ReadBinary(string callbackId)
        {
            if (string.IsNullOrEmpty(currentPageId)) currentPageId = GetCurrentPageId();
            string value;
            application.GetBinaryPageContent(currentPageId, callbackId, out value);
            return value;
        }

        public void UpdateCurrentPage(XDocument page)
        {
            if (string.IsNullOrEmpty(currentPageId)) currentPageId = GetCurrentPageId();
            application.UpdatePageContent(page.ToString(SaveOptions.DisableFormatting), DateTime.MinValue, Schema2013, true);
        }

        private string GetCurrentPageId()
        {
            dynamic windows = null;
            dynamic window = null;
            try
            {
                windows = application.Windows;
                window = windows.CurrentWindow;
                if (window == null || string.IsNullOrEmpty(window.CurrentPageId))
                {
                    throw new InvalidOperationException("OneNote hat keine aktive Seite.");
                }
                return window.CurrentPageId;
            }
            finally
            {
                Release(window);
                Release(windows);
            }
        }

        public void Dispose()
        {
            if (application != null)
            {
                Release(application);
                application = null;
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
