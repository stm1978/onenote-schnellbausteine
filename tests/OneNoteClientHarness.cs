using System;
using System.Collections.Generic;
using System.IO;
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
            public Exception UpdateException;

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
                throw new InvalidOperationException(
                    "Nach dem Update darf die Seite nicht erneut gelesen werden.");
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
                throw new InvalidOperationException(
                    "Der Insert-Pfad darf NavigateTo nicht aufrufen.");
            }

            public void Dispose()
            {
            }
        }

        private sealed class FakeWindowActivator : IOneNoteWindowActivator
        {
            private readonly List<string> calls;
            public Exception ActivationException;

            public FakeWindowActivator(List<string> calls)
            {
                this.calls = calls;
            }

            public void Activate(IntPtr windowHandle)
            {
                calls.Add("activate:" + windowHandle.ToInt64());
                if (ActivationException != null) throw ActivationException;
            }
        }

        private sealed class FakeWindowNative : IOneNoteWindowNative
        {
            public readonly List<string> Calls;
            public IntPtr ForegroundWindow = new IntPtr(200);
            public IntPtr FocusedWindow = new IntPtr(101);
            public bool Iconic;
            public bool SetForegroundResult = true;
            public bool BringToTopResult = true;
            public bool RestoreResult = true;
            public bool ThrowWhenSettingForeground;

            public FakeWindowNative(List<string> calls)
            {
                Calls = calls;
            }

            public bool AttachThreadInput(
                uint attachThread,
                uint attachToThread,
                bool attach)
            {
                Calls.Add(
                    "attach:" + attachThread + ":" + attachToThread + ":" + attach);
                return true;
            }

            public bool BringWindowToTop(IntPtr windowHandle)
            {
                Calls.Add("bring:" + windowHandle.ToInt64());
                return BringToTopResult;
            }

            public uint GetCurrentProcessId()
            {
                return 1;
            }

            public uint GetCurrentThreadId()
            {
                return 10;
            }

            public IntPtr GetFocus()
            {
                Calls.Add("focus");
                return FocusedWindow;
            }

            public IntPtr GetForegroundWindow()
            {
                Calls.Add("foreground:" + ForegroundWindow.ToInt64());
                return ForegroundWindow;
            }

            public uint GetWindowThreadProcessId(
                IntPtr windowHandle,
                out uint processId)
            {
                if (windowHandle == new IntPtr(100))
                {
                    processId = 3;
                    return 30;
                }
                if (windowHandle == new IntPtr(200))
                {
                    processId = 2;
                    return 20;
                }
                processId = 0;
                return 0;
            }

            public bool IsChild(IntPtr parentWindow, IntPtr windowHandle)
            {
                Calls.Add(
                    "child:" + parentWindow.ToInt64() + ":" + windowHandle.ToInt64());
                return parentWindow == new IntPtr(100) &&
                    windowHandle == new IntPtr(101);
            }

            public bool IsIconic(IntPtr windowHandle)
            {
                Calls.Add("iconic:" + windowHandle.ToInt64());
                return Iconic;
            }

            public bool IsWindow(IntPtr windowHandle)
            {
                Calls.Add("window:" + windowHandle.ToInt64());
                return windowHandle == new IntPtr(100);
            }

            public bool SetForegroundWindow(IntPtr windowHandle)
            {
                Calls.Add("set:" + windowHandle.ToInt64());
                if (ThrowWhenSettingForeground)
                {
                    throw new InvalidOperationException("set failed");
                }
                if (SetForegroundResult) ForegroundWindow = windowHandle;
                return SetForegroundResult;
            }

            public bool ShowWindow(IntPtr windowHandle, int command)
            {
                Calls.Add("show:" + windowHandle.ToInt64() + ":" + command);
                return RestoreResult;
            }
        }

        private static int Main()
        {
            try
            {
                InsertNeverNavigatesAndActivationFollowsUpdate();
                ActivationFailureDoesNotFailSuccessfulUpdate();
                UpdateFailureDoesNotAttemptActivation();
                GalleryItemsAreSortedAndAddressable();
                EmptyGalleryItemIsDisabled();
                ActivationFailureIsNonModalAndDoesNotFailInsert();
                GalleryInvalidationReloadsCrud();
                TargetAlreadyForegroundAcceptsSetForegroundFalse();
                CurrentForegroundAndTargetThreadsAreAttached();
                AttachedThreadsAreDetachedInFinally();
                MinimizedTargetIsRestoredBeforeActivation();
                RejectedActivationDoesNotFailInsert();
                Console.WriteLine("12/12 OneNote-Client-Harness-Tests erfolgreich.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }

        private static void InsertNeverNavigatesAndActivationFollowsUpdate()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            var activator = new FakeWindowActivator(calls);
            using (var client = CreateClient(application, activator))
            {
                InsertSnippetOperation.Execute(
                    client,
                    Page(),
                    new Snippet { Id = "snippet-id", Name = "Test" },
                    delegate { return Fragment(); });
            }

            Equal(
                "update,activate:123",
                string.Join(",", calls));
            NoCallStartsWith(calls, "navigate:");
            NoCallStartsWith(calls, "read:");
        }

        private static void ActivationFailureDoesNotFailSuccessfulUpdate()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            var activator = new FakeWindowActivator(calls)
            {
                ActivationException = new InvalidOperationException("activation failed")
            };
            using (var client = CreateClient(application, activator))
            {
                client.UpdateCurrentPageAndActivateWindow(Page());
            }
            Equal("update,activate:123", string.Join(",", calls));
        }

        private static void UpdateFailureDoesNotAttemptActivation()
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
                    delegate { client.UpdateCurrentPageAndActivateWindow(Page()); });
            }
            Equal("update", string.Join(",", calls));
        }

        private static void GalleryItemsAreSortedAndAddressable()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                Snippet zulu = repository.Save("Zulu", Fragment(), false);
                Snippet alpha = repository.Save("Alpha", Fragment(), false);
                var gallery = new SnippetGalleryModel(repository);

                Equal(2, gallery.RefreshAndGetCount());
                Equal("Alpha", gallery.GetLabel(0));
                Equal("Zulu", gallery.GetLabel(1));
                Equal("sqSnippet_" + alpha.Id, gallery.GetId(0));
                Equal("sqSnippet_" + zulu.Id, gallery.GetId(1));
                Equal(true, gallery.RefreshAndGetEnabled());
                Equal(alpha.Id, gallery.GetSnippet("sqSnippet_" + alpha.Id, 0).Id);
                Throws<InvalidOperationException>(delegate
                {
                    gallery.GetSnippet("sqSnippet_" + zulu.Id, 0);
                });
            });
        }

        private static void EmptyGalleryItemIsDisabled()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var gallery = new SnippetGalleryModel(repository);
                Equal(1, gallery.RefreshAndGetCount());
                Equal("Keine Schnellbausteine gespeichert", gallery.GetLabel(0));
                Equal("sqEmpty", gallery.GetId(0));
                Equal(false, gallery.RefreshAndGetEnabled());
                Throws<InvalidOperationException>(
                    delegate { gallery.GetSnippet("sqEmpty", 0); });
            });
        }

        private static void ActivationFailureIsNonModalAndDoesNotFailInsert()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            var activator = new FakeWindowActivator(calls)
            {
                ActivationException = new InvalidOperationException("activation failed")
            };
            using (var client = CreateClient(application, activator))
            {
                InsertSnippetOperation.Execute(
                    client,
                    Page(),
                    new Snippet { Id = "snippet-id", Name = "Test" },
                    delegate { return Fragment(); });
            }
            Equal(
                "update,activate:123",
                string.Join(",", calls));
            NoCallStartsWith(calls, "navigate:");
        }

        private static void GalleryInvalidationReloadsCrud()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                Snippet alpha = repository.Save("Alpha", Fragment(), false);
                var gallery = new SnippetGalleryModel(repository);
                Equal(1, gallery.RefreshAndGetCount());

                Snippet zulu = repository.Save("Zulu", Fragment(), false);
                gallery.Invalidate();
                Throws<InvalidOperationException>(delegate { gallery.GetLabel(0); });
                Equal(2, gallery.RefreshAndGetCount());

                repository.Rename(zulu.Id, "Aaron");
                gallery.Invalidate();
                Equal(2, gallery.RefreshAndGetCount());
                Equal("Aaron", gallery.GetLabel(0));

                repository.Delete(alpha.Id);
                repository.Delete(zulu.Id);
                gallery.Invalidate();
                Equal(1, gallery.RefreshAndGetCount());
                Equal(false, gallery.RefreshAndGetEnabled());
            });
        }

        private static void TargetAlreadyForegroundAcceptsSetForegroundFalse()
        {
            var calls = new List<string>();
            var native = new FakeWindowNative(calls)
            {
                ForegroundWindow = new IntPtr(100),
                SetForegroundResult = false
            };

            new OneNoteWindowActivator(native).Activate(new IntPtr(100));

            Ordered(
                calls,
                "attach:10:30:True",
                "set:100",
                "bring:100",
                "focus",
                "child:100:101",
                "attach:10:30:False");
        }

        private static void CurrentForegroundAndTargetThreadsAreAttached()
        {
            var calls = new List<string>();
            var native = new FakeWindowNative(calls);

            new OneNoteWindowActivator(native).Activate(new IntPtr(100));

            Ordered(
                calls,
                "attach:10:20:True",
                "attach:10:30:True",
                "set:100",
                "bring:100",
                "attach:10:30:False",
                "attach:10:20:False");
        }

        private static void AttachedThreadsAreDetachedInFinally()
        {
            var calls = new List<string>();
            var native = new FakeWindowNative(calls)
            {
                ThrowWhenSettingForeground = true
            };

            Throws<InvalidOperationException>(
                delegate
                {
                    new OneNoteWindowActivator(native).Activate(new IntPtr(100));
                });

            Ordered(
                calls,
                "attach:10:20:True",
                "attach:10:30:True",
                "set:100",
                "attach:10:30:False",
                "attach:10:20:False");
        }

        private static void MinimizedTargetIsRestoredBeforeActivation()
        {
            var calls = new List<string>();
            var native = new FakeWindowNative(calls)
            {
                Iconic = true
            };

            new OneNoteWindowActivator(native).Activate(new IntPtr(100));

            Ordered(calls, "iconic:100", "show:100:9", "set:100", "bring:100");
        }

        private static void RejectedActivationDoesNotFailInsert()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            var native = new FakeWindowNative(calls)
            {
                SetForegroundResult = false,
                BringToTopResult = false,
                FocusedWindow = IntPtr.Zero
            };
            using (var client = new OneNoteClient(
                application,
                "page-id",
                new IntPtr(100),
                new DateTime(2026, 8, 8, 10, 0, 0, DateTimeKind.Utc),
                new OneNoteWindowActivator(native)))
            {
                InsertSnippetOperation.Execute(
                    client,
                    Page(),
                    new Snippet { Id = "snippet-id", Name = "Test" },
                    delegate { return Fragment(); });
            }
            Contains(calls, "set:100");
            Contains(calls, "bring:100");
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

        private static void WithTemporaryRepository(Action<SnippetRepository> action)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "Schnellbausteine.OneNoteClientHarness",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            try
            {
                action(new SnippetRepository(path));
            }
            finally
            {
                Directory.Delete(path, true);
            }
        }

        private static void Equal<T>(T expected, T actual)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
            {
                throw new Exception("Erwartet: " + expected + "; tatsächlich: " + actual);
            }
        }

        private static void Contains(List<string> values, string expected)
        {
            if (!values.Contains(expected))
            {
                throw new Exception("Aufruf fehlt: " + expected);
            }
        }

        private static void NoCallStartsWith(List<string> values, string prefix)
        {
            foreach (string value in values)
            {
                if (value.StartsWith(prefix, StringComparison.Ordinal))
                {
                    throw new Exception("Verbotener Aufruf: " + value);
                }
            }
        }

        private static void Ordered(List<string> values, params string[] expected)
        {
            int index = -1;
            foreach (string value in expected)
            {
                index = values.IndexOf(value, index + 1);
                if (index < 0)
                {
                    throw new Exception(
                        "Aufruf fehlt oder ist falsch angeordnet: " + value +
                        "; tatsächlich: " + string.Join(",", values));
                }
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
