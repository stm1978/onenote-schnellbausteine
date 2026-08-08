using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Schnellbausteine.AddIn;
using Schnellbausteine.Core;

namespace Schnellbausteine.Tests
{
    internal static class OneNoteClientHarness
    {
        private sealed class FakeApplication : IOneNoteApplicationAdapter
        {
            public readonly List<string> Calls = new List<string>();
            public string RefreshedXml;
            public Exception UpdateException;
            public Exception NavigateException;
            public string NavigatedPageId;
            public string NavigatedObjectId;

            public OneNoteWindowInfo GetCurrentWindow()
            {
                throw new NotSupportedException();
            }

            public string GetPageContent(
                string pageId,
                bool includeBinaryData,
                bool includeSelection)
            {
                Calls.Add(
                    "read:" + pageId + ":" + includeBinaryData + ":" + includeSelection);
                return RefreshedXml;
            }

            public string GetBinaryPageContent(string pageId, string callbackId)
            {
                throw new NotSupportedException();
            }

            public void UpdatePageContent(string xml, DateTime expectedLastModifiedTime)
            {
                Calls.Add("update");
                if (UpdateException != null) throw UpdateException;
            }

            public void NavigateTo(string pageId, string objectId)
            {
                Calls.Add("navigate:" + pageId + ":" + objectId);
                NavigatedPageId = pageId;
                NavigatedObjectId = objectId;
                if (NavigateException != null) throw NavigateException;
            }

            public void Dispose()
            {
            }
        }

        private static int Main()
        {
            try
            {
                UpdateReadBackAndNavigateAreOrdered();
                MissingFocusTargetIsReportedAfterUpdate();
                NavigateFailureIsReportedAfterUpdate();
                UpdateFailureDoesNotAttemptFocus();
                Console.WriteLine("4/4 OneNote-Client-Harness-Tests erfolgreich.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }

        private static void UpdateReadBackAndNavigateAreOrdered()
        {
            var application = new FakeApplication
            {
                RefreshedXml = PageXml("focus-object")
            };
            using (var client = CreateClient(application))
            {
                client.UpdateCurrentPageAndRestoreFocus(Page());
            }

            Equal(
                "update,read:page-id:False:True,navigate:page-id:focus-object",
                string.Join(",", application.Calls));
            Equal("page-id", application.NavigatedPageId);
            Equal("focus-object", application.NavigatedObjectId);
        }

        private static void MissingFocusTargetIsReportedAfterUpdate()
        {
            var application = new FakeApplication
            {
                RefreshedXml = PageXml(null)
            };
            using (var client = CreateClient(application))
            {
                Throws<OneNoteFocusException>(
                    delegate { client.UpdateCurrentPageAndRestoreFocus(Page()); });
            }
            Equal("update,read:page-id:False:True", string.Join(",", application.Calls));
        }

        private static void NavigateFailureIsReportedAfterUpdate()
        {
            var application = new FakeApplication
            {
                RefreshedXml = PageXml("focus-object"),
                NavigateException = new COMException("navigation failed")
            };
            using (var client = CreateClient(application))
            {
                Throws<OneNoteFocusException>(
                    delegate { client.UpdateCurrentPageAndRestoreFocus(Page()); });
            }
            Equal(
                "update,read:page-id:False:True,navigate:page-id:focus-object",
                string.Join(",", application.Calls));
        }

        private static void UpdateFailureDoesNotAttemptFocus()
        {
            var application = new FakeApplication
            {
                UpdateException = new COMException("update failed")
            };
            using (var client = CreateClient(application))
            {
                Throws<COMException>(
                    delegate { client.UpdateCurrentPageAndRestoreFocus(Page()); });
            }
            Equal("update", string.Join(",", application.Calls));
        }

        private static OneNoteClient CreateClient(FakeApplication application)
        {
            return new OneNoteClient(
                application,
                "page-id",
                new IntPtr(123),
                new DateTime(2026, 8, 8, 10, 0, 0, DateTimeKind.Utc));
        }

        private static XDocument Page()
        {
            return XDocument.Parse(
                "<one:Page xmlns:one=\"" + OneNoteXml.One + "\">" +
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline></one:Page>",
                LoadOptions.PreserveWhitespace);
        }

        private static string PageXml(string objectId)
        {
            string id = objectId == null ? string.Empty : " objectID=\"" + objectId + "\"";
            return "<one:Page xmlns:one=\"" + OneNoteXml.One + "\">" +
                "<one:Outline selected=\"partial\"><one:OEChildren><one:OE" + id +
                " selected=\"partial\">" +
                "<one:T selected=\"all\"><![CDATA[]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline></one:Page>";
        }

        private static void Equal<T>(T expected, T actual)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
            {
                throw new Exception("Erwartet: " + expected + "; tatsächlich: " + actual);
            }
        }

        private static void Throws<T>(Action action) where T : Exception
        {
            try
            {
                action();
            }
            catch (T)
            {
                return;
            }
            throw new Exception("Ausnahme " + typeof(T).Name + " wurde erwartet.");
        }
    }
}
