using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using Schnellbausteine.AddIn;
using Schnellbausteine.Core;

namespace Schnellbausteine.Tests
{
    internal static class OneNoteClientHarness
    {
        private sealed class FakeApplication : IOneNoteApplicationAdapter
        {
            public readonly List<string> Calls;
            public string RefreshedXml;
            public Exception UpdateException;
            public Exception NavigateException;
            public string NavigatedPageId;
            public string NavigatedObjectId;

            public FakeApplication(List<string> calls)
            {
                Calls = calls;
            }

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

        private sealed class FakeWindowActivator : IOneNoteWindowActivator
        {
            private readonly List<string> calls;

            public FakeWindowActivator(List<string> calls)
            {
                this.calls = calls;
            }

            public OneNoteWindowActivationException Failure;

            public void Activate(IntPtr windowHandle)
            {
                calls.Add("activate:" + windowHandle.ToInt64());
                if (Failure != null) throw Failure;
            }
        }

        private sealed class FakePicker : IInsertSnippetPicker
        {
            private readonly List<string> calls;
            private readonly bool accepted;

            public FakePicker(List<string> calls, bool accepted)
            {
                this.calls = calls;
                this.accepted = accepted;
            }

            public bool TrySelect(IWin32Window owner, out Snippet snippet)
            {
                calls.Add("picker.show");
                snippet = accepted ? new Snippet { Id = "snippet-id", Name = "Test" } : null;
                return accepted;
            }

            public void Dispose()
            {
                calls.Add("picker.dispose");
            }
        }

        private static int Main()
        {
            try
            {
                UpdateReadBackNavigateAndActivationAreOrdered();
                MissingFocusTargetIsReportedAfterUpdate();
                NavigateFailureIsReportedAfterUpdate();
                ActivationFailureIsReportedAfterNavigate();
                UpdateFailureDoesNotAttemptFocus();
                PickerIsDisposedBeforeUpdateAndFocus();
                FocusFailureIsReportedWithoutFailingInsert();
                CancelledPickerDoesNotUpdate();
                Console.WriteLine("8/8 OneNote-Client-Harness-Tests erfolgreich.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }

        private static void UpdateReadBackNavigateAndActivationAreOrdered()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                RefreshedXml = PageXml("focus-object")
            };
            var activator = new FakeWindowActivator(calls);
            using (var client = CreateClient(application, activator))
            {
                client.UpdateCurrentPageAndRestoreFocus(Page());
            }

            Equal(
                "update,read:page-id:False:True,navigate:page-id:focus-object,activate:123",
                string.Join(",", calls));
            Equal("page-id", application.NavigatedPageId);
            Equal("focus-object", application.NavigatedObjectId);
        }

        private static void MissingFocusTargetIsReportedAfterUpdate()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                RefreshedXml = PageXml(null)
            };
            using (var client = CreateClient(
                application,
                new FakeWindowActivator(calls)))
            {
                Throws<OneNoteFocusException>(
                    delegate { client.UpdateCurrentPageAndRestoreFocus(Page()); });
            }
            Equal("update,read:page-id:False:True", string.Join(",", calls));
        }

        private static void NavigateFailureIsReportedAfterUpdate()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                RefreshedXml = PageXml("focus-object"),
                NavigateException = new COMException("navigation failed")
            };
            using (var client = CreateClient(
                application,
                new FakeWindowActivator(calls)))
            {
                Throws<OneNoteFocusException>(
                    delegate { client.UpdateCurrentPageAndRestoreFocus(Page()); });
            }
            Equal(
                "update,read:page-id:False:True,navigate:page-id:focus-object",
                string.Join(",", calls));
        }

        private static void ActivationFailureIsReportedAfterNavigate()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                RefreshedXml = PageXml("focus-object")
            };
            var activator = new FakeWindowActivator(calls)
            {
                Failure = new OneNoteWindowActivationException("activation failed")
            };
            using (var client = CreateClient(application, activator))
            {
                Throws<OneNoteFocusException>(
                    delegate { client.UpdateCurrentPageAndRestoreFocus(Page()); });
            }
            Equal(
                "update,read:page-id:False:True,navigate:page-id:focus-object,activate:123",
                string.Join(",", calls));
        }

        private static void UpdateFailureDoesNotAttemptFocus()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                UpdateException = new COMException("update failed")
            };
            using (var client = CreateClient(
                application,
                new FakeWindowActivator(calls)))
            {
                Throws<COMException>(
                    delegate { client.UpdateCurrentPageAndRestoreFocus(Page()); });
            }
            Equal("update", string.Join(",", calls));
        }

        private static void PickerIsDisposedBeforeUpdateAndFocus()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                RefreshedXml = PageXml("focus-object")
            };
            using (var client = CreateClient(
                application,
                new FakeWindowActivator(calls)))
            {
                bool inserted = InsertSnippetOperation.Execute(
                    client,
                    Page(),
                    new WindowHandleOwner(new IntPtr(123)),
                    delegate { return new FakePicker(calls, true); },
                    delegate { return Fragment(); },
                    delegate { throw new Exception("Kein Fokusfehler erwartet."); });
                Equal(true, inserted);
            }
            Equal(
                "picker.show,picker.dispose,update,read:page-id:False:True," +
                "navigate:page-id:focus-object,activate:123",
                string.Join(",", calls));
        }

        private static void FocusFailureIsReportedWithoutFailingInsert()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                RefreshedXml = PageXml("focus-object")
            };
            var activator = new FakeWindowActivator(calls)
            {
                Failure = new OneNoteWindowActivationException("activation failed")
            };
            using (var client = CreateClient(application, activator))
            {
                bool inserted = InsertSnippetOperation.Execute(
                    client,
                    Page(),
                    new WindowHandleOwner(new IntPtr(123)),
                    delegate { return new FakePicker(calls, true); },
                    delegate { return Fragment(); },
                    delegate { calls.Add("focus.reported"); });
                Equal(true, inserted);
            }
            Equal(
                "picker.show,picker.dispose,update,read:page-id:False:True," +
                "navigate:page-id:focus-object,activate:123,focus.reported",
                string.Join(",", calls));
        }

        private static void CancelledPickerDoesNotUpdate()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            using (var client = CreateClient(
                application,
                new FakeWindowActivator(calls)))
            {
                bool inserted = InsertSnippetOperation.Execute(
                    client,
                    Page(),
                    new WindowHandleOwner(new IntPtr(123)),
                    delegate { return new FakePicker(calls, false); },
                    delegate { throw new Exception("Fragment darf nicht geladen werden."); },
                    delegate { throw new Exception("Kein Fokusfehler erwartet."); });
                Equal(false, inserted);
            }
            Equal("picker.show,picker.dispose", string.Join(",", calls));
        }

        private static OneNoteClient CreateClient(
            FakeApplication application,
            IOneNoteWindowActivator activator)
        {
            return new OneNoteClient(
                application,
                "page-id",
                new IntPtr(123),
                new DateTime(2026, 8, 8, 10, 0, 0, DateTimeKind.Utc),
                activator);
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

        private static XDocument Fragment()
        {
            return XDocument.Parse(
                "<sq:Snippet xmlns:sq=\"" + OneNoteXml.Snippet +
                "\" xmlns:one=\"" + OneNoteXml.One + "\" version=\"1\">" +
                "<sq:TagDefs/><sq:QuickStyleDefs/><sq:Content>" +
                "<one:OE><one:T><![CDATA[snippet]]></one:T></one:OE>" +
                "</sq:Content></sq:Snippet>",
                LoadOptions.PreserveWhitespace);
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
