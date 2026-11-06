using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Schnellbausteine.AddIn;
using Schnellbausteine.Core;

namespace Schnellbausteine.Tests
{
    internal static class OneNoteClientHarness
    {
        private sealed class FakeRibbonControl : Microsoft.Office.Core.IRibbonControl
        {
            private readonly string id;

            public FakeRibbonControl(string id)
            {
                this.id = id;
            }

            public string Id { get { return id; } }
            public object Context { get { return null; } }
            public string Tag { get { return string.Empty; } }
        }

        private sealed class FakeApplication : IOneNoteApplicationAdapter
        {
            public readonly List<string> Calls;
            public Exception UpdateException;
            public string BinaryData;
            public string PageXml;
            public string UpdatedXml;

            public FakeApplication(List<string> calls)
            {
                Calls = calls;
            }

            public OneNoteWindowInfo GetCurrentWindow()
            {
                if (PageXml != null)
                {
                    Calls.Add("window");
                    return new OneNoteWindowInfo { PageId = "capture-page", WindowHandle = new IntPtr(123) };
                }
                throw new NotSupportedException();
            }

            public string GetPageContent(
                string pageId,
                bool includeBinaryData,
                bool includeSelection)
            {
                Calls.Add(
                    "read:" + pageId + ":" + includeBinaryData + ":" + includeSelection);
                if (PageXml != null) return PageXml;
                throw new InvalidOperationException(
                    "Nach dem Update darf die Seite nicht erneut gelesen werden.");
            }

            public string GetBinaryPageContent(string pageId, string callbackId)
            {
                Calls.Add("binary:" + pageId + ":" + callbackId);
                return BinaryData;
            }

            public void UpdatePageContent(string xml, DateTime expectedLastModifiedTime)
            {
                Calls.Add("update");
                UpdatedXml = xml;
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
            public IntPtr RootWindow = new IntPtr(100);
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

            public IntPtr GetRootWindow(IntPtr windowHandle)
            {
                Calls.Add("root:" + windowHandle.ToInt64());
                return RootWindow;
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
                return windowHandle == new IntPtr(100) || windowHandle == new IntPtr(101);
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
                MenuItemsAreSortedEscapedAndAddressable();
                EmptyMenuItemIsDisabled();
                ActivationFailureIsNonModalAndDoesNotFailInsert();
                MenuInvalidationRejectsStaleIdsAndReloadsCrud();
                TargetAlreadyForegroundAcceptsSetForegroundFalse();
                CurrentForegroundAndTargetThreadsAreAttached();
                WorkspaceChildActivatesTheRootFrame();
                MissingRootNeverActivatesAWindow();
                InvalidRootNeverActivatesAWindow();
                InvalidWorkspaceNeverResolvesOrActivatesAWindow();
                AttachedThreadsAreDetachedInFinally();
                MinimizedTargetIsRestoredBeforeActivation();
                RejectedActivationDoesNotFailInsert();
                BinaryCallbacksUseCapturedPageOnClientThread();
                CaptureReadDefersBinaryAndPreservesSelectedPage();
                InsertReadStillIncludesPageBinaryData();
                InkInTargetContainerBlocksInsertWithoutMutation();
                InkSnippetBlocksInsertWithoutMutation();
                DirectInkUpdateIsBlockedBeforeCom();
                TextOnlyInsertStillUpdatesThePage();
                ScopedInsertOmitsUnrelatedInkAndPageObjects();
                ScopedInsertRejectsChangesOutsideTarget();
                ScopedInsertIncludesOnlyNewDefinitions();
                UnidentifiedTargetNeverUpdates();
                ScopedBlockInsertRetainsCompleteTargetContent();
                FailedScopedUpdateLeavesSourceUnchanged();
                ExpirationSurfaceIsRemoved();
                FeatureActionsRemainAvailableWithoutClock();
                Console.WriteLine("31/31 OneNote-Client-Harness-Tests erfolgreich.");
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

        private static void MenuItemsAreSortedEscapedAndAddressable()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                Snippet zulu = repository.Save("Zulu", Fragment(), false);
                Snippet alpha = repository.Save("A & <Alpha>", Fragment(), false);
                var menu = new SnippetMenuModel(repository);
                string contentXml = menu.RefreshAndGetContent("sqMain");
                XDocument content = XDocument.Parse(contentXml);
                XNamespace ribbon =
                    "http://schemas.microsoft.com/office/2009/07/customui";
                XElement[] children = content.Root.Elements().ToArray();

                Equal(4, children.Length);
                Equal("button,button,menuSeparator,button",
                    string.Join(",", children.Select(item => item.Name.LocalName)));
                Equal("A & <Alpha>", (string)children[0].Attribute("label"));
                Equal(true, contentXml.Contains(
                    "label=\"A &amp; &lt;Alpha&gt;\""));
                Equal("Zulu", (string)children[1].Attribute("label"));
                Equal("Schnellbausteine verwalten",
                    (string)children[3].Attribute("label"));
                Equal("ManageSnippets", (string)children[3].Attribute("onAction"));
                Equal(alpha.Id, menu.GetSnippet(
                    (string)children[0].Attribute("id")).Id);
                Equal(zulu.Id, menu.GetSnippet(
                    (string)children[1].Attribute("id")).Id);
                Equal(true, Regex.IsMatch(
                    (string)children[0].Attribute("id"),
                    "^sqSnippet_[0-9a-f]{32}_0$"));
                Equal(false, ((string)children[0].Attribute("id"))
                    .Contains(alpha.Id));
                Equal(true, children.All(item => item.Name.Namespace == ribbon));
                Throws<InvalidOperationException>(
                    delegate { menu.GetSnippet("sqSnippet_" + alpha.Id); });
            });
        }

        private static void EmptyMenuItemIsDisabled()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var menu = new SnippetMenuModel(repository);
                XDocument content = XDocument.Parse(
                    menu.RefreshAndGetContent("sqInsertText"));
                XElement[] children = content.Root.Elements().ToArray();
                Equal(3, children.Length);
                Equal("Keine Schnellbausteine gespeichert",
                    (string)children[0].Attribute("label"));
                Equal("false", (string)children[0].Attribute("enabled"));
                Equal("menuSeparator", children[1].Name.LocalName);
                Equal("Schnellbausteine verwalten",
                    (string)children[2].Attribute("label"));
                Throws<InvalidOperationException>(
                    delegate { menu.GetSnippet("sqEmpty"); });
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

        private static void MenuInvalidationRejectsStaleIdsAndReloadsCrud()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                Snippet alpha = repository.Save("Alpha", Fragment(), false);
                var menu = new SnippetMenuModel(repository);
                XDocument first = XDocument.Parse(
                    menu.RefreshAndGetContent("sqMain"));
                string staleId = (string)first.Root.Elements().First()
                    .Attribute("id");
                XDocument context = XDocument.Parse(
                    menu.RefreshAndGetContent("sqInsertText"));
                string contextId = (string)context.Root.Elements().First()
                    .Attribute("id");
                Equal(alpha.Id, menu.GetSnippet(staleId).Id);
                Equal(alpha.Id, menu.GetSnippet(contextId).Id);

                XDocument refreshedRibbon = XDocument.Parse(
                    menu.RefreshAndGetContent("sqMain"));
                string refreshedRibbonId = (string)refreshedRibbon.Root
                    .Elements().First().Attribute("id");
                Throws<InvalidOperationException>(
                    delegate { menu.GetSnippet(staleId); });
                Equal(alpha.Id, menu.GetSnippet(contextId).Id);
                Equal(alpha.Id, menu.GetSnippet(refreshedRibbonId).Id);

                Snippet zulu = repository.Save("Zulu", Fragment(), false);
                menu.Invalidate();
                Throws<InvalidOperationException>(
                    delegate { menu.GetSnippet(contextId); });
                XDocument second = XDocument.Parse(
                    menu.RefreshAndGetContent("sqMain"));
                Equal(4, second.Root.Elements().Count());
                Throws<InvalidOperationException>(
                    delegate { menu.GetSnippet(contextId); });

                repository.Rename(zulu.Id, "Aaron");
                menu.Invalidate();
                XDocument third = XDocument.Parse(
                    menu.RefreshAndGetContent("sqInsertOutline"));
                Equal("Aaron", (string)third.Root.Elements().First()
                    .Attribute("label"));

                repository.Delete(alpha.Id);
                repository.Delete(zulu.Id);
                menu.Invalidate();
                XDocument empty = XDocument.Parse(
                    menu.RefreshAndGetContent("sqInsertInk"));
                Equal("false", (string)empty.Root.Elements().First()
                    .Attribute("enabled"));
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

        private static void WorkspaceChildActivatesTheRootFrame()
        {
            var calls = new List<string>();
            var native = new FakeWindowNative(calls);

            new OneNoteWindowActivator(native).Activate(new IntPtr(101));

            Ordered(calls, "root:101", "set:100", "bring:100", "child:100:101");
            Equal(false, calls.Contains("set:101"));
            Equal(false, calls.Contains("bring:101"));
        }

        private static void MissingRootNeverActivatesAWindow()
        {
            var calls = new List<string>();
            var native = new FakeWindowNative(calls) { RootWindow = IntPtr.Zero };

            new OneNoteWindowActivator(native).Activate(new IntPtr(101));

            Contains(calls, "root:101");
            Equal(false, calls.Any(call => call.StartsWith("set:", StringComparison.Ordinal)));
            Equal(false, calls.Any(call => call.StartsWith("attach:", StringComparison.Ordinal)));
        }

        private static void InvalidRootNeverActivatesAWindow()
        {
            var calls = new List<string>();
            var native = new FakeWindowNative(calls) { RootWindow = new IntPtr(999) };

            new OneNoteWindowActivator(native).Activate(new IntPtr(101));

            Contains(calls, "root:101");
            Equal(false, calls.Any(call => call.StartsWith("set:", StringComparison.Ordinal)));
            Equal(false, calls.Any(call => call.StartsWith("attach:", StringComparison.Ordinal)));
        }

        private static void InvalidWorkspaceNeverResolvesOrActivatesAWindow()
        {
            foreach (IntPtr handle in new[] { IntPtr.Zero, new IntPtr(999) })
            {
                var calls = new List<string>();
                new OneNoteWindowActivator(new FakeWindowNative(calls)).Activate(handle);
                Equal(false, calls.Any(call => call.StartsWith("root:", StringComparison.Ordinal)));
                Equal(false, calls.Any(call => call.StartsWith("set:", StringComparison.Ordinal)));
                Equal(false, calls.Any(call => call.StartsWith("attach:", StringComparison.Ordinal)));
            }
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

        private static void BinaryCallbacksUseCapturedPageOnClientThread()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls) { BinaryData = "opaque-data" };
            using (var client = CreateClient(
                application,
                new FakeWindowActivator(calls)))
            {
                Equal("opaque-data", client.ReadBinary("ink-callback"));
            }
            Equal("binary:page-id:ink-callback", string.Join(",", calls));
        }

        private static void CaptureReadDefersBinaryAndPreservesSelectedPage()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                PageXml = "<one:Page xmlns:one=\"" + OneNoteXml.One +
                    "\" lastModifiedTime=\"2026-11-06T10:00:00Z\">" +
                    "<one:InkDrawing selected=\"all\" objectID=\"ink-object\"/>" +
                    "<one:InkDrawing objectID=\"unselected-ink\"/>" +
                    "</one:Page>",
                BinaryData = "AQID"
            };
            using (var client = new OneNoteClient(application, new FakeWindowActivator(calls)))
            {
                XDocument capturedPage = client.ReadCurrentPageForCapture();
                Equal("window,read:capture-page:False:True", string.Join(",", calls));
                // Changing focus while naming must not trigger another page/selection read.
                application.PageXml = Page().ToString();
                XDocument fragment = OneNoteXml.ExtractSelection(capturedPage, client.ReadBinary);
                Equal("AQID", fragment.Descendants(OneNoteXml.One + "Data").Single().Value);
            }
            Equal("window,read:capture-page:False:True,binary:capture-page:ink-object",
                string.Join(",", calls));
        }

        private static void InsertReadStillIncludesPageBinaryData()
        {
            var calls = new List<string>();
            XDocument page = Page();
            page.Root.SetAttributeValue("lastModifiedTime", "2026-11-06T10:00:00Z");
            var application = new FakeApplication(calls) { PageXml = page.ToString() };
            using (var client = new OneNoteClient(application, new FakeWindowActivator(calls)))
            {
                client.ReadCurrentPage();
            }
            Equal("window,read:capture-page:True:True", string.Join(",", calls));
        }

        private static void InkInTargetContainerBlocksInsertWithoutMutation()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            XDocument page = Page();
            var ink = new XElement(OneNoteXml.One + "InkDrawing",
                new XAttribute("objectID", "original-ink"),
                new XElement(OneNoteXml.One + "Data", "AQID"));
            page.Descendants(OneNoteXml.One + "OE").Single().Add(ink);
            string original = page.ToString();
            bool loaded = false;
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                Throws<InvalidOperationException>(delegate
                {
                    InsertSnippetOperation.Execute(
                        client, page, new Snippet { Id = "text" },
                        delegate
                        {
                            loaded = true;
                            return Fragment();
                        });
                });
            }
            Equal(false, loaded);
            Equal(original, page.ToString());
            Equal(0, calls.Count);
            Equal(null, application.UpdatedXml);
        }

        private static void InkSnippetBlocksInsertWithoutMutation()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            XDocument page = Page();
            string original = page.ToString();
            XDocument fragment = Fragment();
            fragment.Descendants(OneNoteXml.One + "OE").Single().ReplaceNodes(
                new XElement(OneNoteXml.One + "InkDrawing",
                    new XElement(OneNoteXml.One + "Data", "AQID")));
            string originalFragment = fragment.ToString();
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                Throws<InvalidOperationException>(delegate
                {
                    InsertSnippetOperation.Execute(
                        client, page, new Snippet { Id = "ink" },
                        delegate { return fragment; });
                });
            }
            Equal(original, page.ToString());
            Equal(originalFragment, fragment.ToString());
            Equal(0, calls.Count);
            Equal(null, application.UpdatedXml);
        }

        private static void DirectInkUpdateIsBlockedBeforeCom()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            XDocument page = Page();
            page.Root.Add(new XElement(OneNoteXml.One + "InkDrawing"));
            string original = page.ToString();
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                Throws<InvalidOperationException>(delegate
                {
                    client.UpdateCurrentPageAndActivateWindow(page);
                });
            }
            Equal(original, page.ToString());
            Equal(0, calls.Count);
            Equal(null, application.UpdatedXml);
        }

        private static void TextOnlyInsertStillUpdatesThePage()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                InsertSnippetOperation.Execute(
                    client, Page(), new Snippet { Id = "text" },
                    delegate { return Fragment(); });
            }
            Equal("update,activate:123", string.Join(",", calls));
            XDocument updated = XDocument.Parse(application.UpdatedXml);
            Equal(true, updated.Descendants(OneNoteXml.One + "T")
                .Any(text => text.Value == "snippet"));
            OneNoteXml.ValidatePostInsertCaret(updated);
        }

        private static void ScopedInsertOmitsUnrelatedInkAndPageObjects()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            XDocument page = Page();
            XNamespace one = OneNoteXml.One;
            page.Root.Add(
                new XElement(one + "InkDrawing", new XAttribute("objectID", "original-ink"),
                    new XElement(one + "Data", "AQID")),
                new XElement(one + "Outline", new XAttribute("objectID", "other-outline"),
                    new XElement(one + "OEChildren", new XElement(one + "OE",
                        new XElement(one + "InkDrawing", new XElement(one + "Data", "BAUG"))))),
                new XElement(one + "Image", new XAttribute("objectID", "original-image")),
                new XElement(one + "Title", new XElement(one + "OE",
                    new XElement(one + "T", "Original title"))),
                new XElement(one + "UnknownFutureData", "must not roundtrip"));
            string original = page.ToString();
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                InsertSnippetOperation.Execute(
                    client, page, new Snippet { Id = "text" }, delegate { return Fragment(); });
            }
            Equal(original, page.ToString());
            XDocument update = XDocument.Parse(application.UpdatedXml);
            Equal("page-id", (string)update.Root.Attribute("ID"));
            Equal(1, update.Root.Elements().Count());
            Equal("target-outline", (string)update.Root.Elements().Single().Attribute("objectID"));
            Equal(false, update.Descendants(one + "InkDrawing").Any());
            Equal(false, update.Descendants(one + "Image").Any());
            Equal(false, update.Descendants(one + "Title").Any());
            Equal(false, update.Descendants(one + "UnknownFutureData").Any());
            Equal("update,activate:123", string.Join(",", calls));
            OneNoteXml.ValidatePostInsertCaret(update);
        }

        private static void ScopedInsertRejectsChangesOutsideTarget()
        {
            XNamespace one = OneNoteXml.One;
            XDocument original = Page();
            original.Root.Add(new XElement(one + "Outline",
                new XAttribute("objectID", "other"), new XElement(one + "OEChildren")));
            foreach (string change in new[] { "modify", "remove", "add" })
            {
                XDocument modified = new XDocument(original);
                OneNoteXml.InsertAtSelection(modified, Fragment());
                XElement other = modified.Root.Elements(one + "Outline").Last();
                if (change == "modify") other.SetAttributeValue("style", "changed");
                if (change == "remove") other.Remove();
                if (change == "add") modified.Root.Add(new XElement(one + "Image"));
                Throws<InvalidOperationException>(delegate
                {
                    InsertSnippetOperation.CreateScopedUpdate(
                        original, modified, original.Root.Elements(one + "Outline").First());
                });
            }
        }

        private static void ScopedInsertIncludesOnlyNewDefinitions()
        {
            XNamespace one = OneNoteXml.One;
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            XDocument page = Page();
            page.Root.AddFirst(
                new XElement(one + "TagDef", new XAttribute("index", "0"),
                    new XAttribute("name", "Existing"), new XAttribute("type", "0"),
                    new XAttribute("symbol", "0")),
                new XElement(one + "QuickStyleDef", new XAttribute("index", "0"),
                    new XAttribute("name", "Existing"), new XAttribute("fontSize", "10.0")));
            XDocument fragment = Fragment();
            fragment.Root.Element(OneNoteXml.Snippet + "TagDefs").Add(
                new XElement(one + "TagDef", new XAttribute("index", "0"),
                    new XAttribute("name", "New"), new XAttribute("type", "0"),
                    new XAttribute("symbol", "1")));
            fragment.Root.Element(OneNoteXml.Snippet + "QuickStyleDefs").Add(
                new XElement(one + "QuickStyleDef", new XAttribute("index", "0"),
                    new XAttribute("name", "New"), new XAttribute("fontSize", "18.0")));
            XElement paragraph = fragment.Descendants(one + "OE").Single();
            paragraph.SetAttributeValue("quickStyleIndex", "0");
            paragraph.AddFirst(new XElement(one + "Tag", new XAttribute("index", "0")));
            string original = page.ToString();
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                InsertSnippetOperation.Execute(
                    client, page, new Snippet { Id = "formatted" }, delegate { return fragment; });
            }
            Equal(original, page.ToString());
            XDocument update = XDocument.Parse(application.UpdatedXml);
            Equal(3, update.Root.Elements().Count());
            Equal("New", (string)update.Root.Element(one + "TagDef").Attribute("name"));
            Equal("New", (string)update.Root.Element(one + "QuickStyleDef").Attribute("name"));
            Equal("1", (string)update.Root.Element(one + "TagDef").Attribute("index"));
            Equal("1", (string)update.Descendants(one + "Tag").Single().Attribute("index"));
            Equal("1", (string)update.Descendants(one + "OE")
                .First(element => element.Attribute("quickStyleIndex") != null)
                .Attribute("quickStyleIndex"));
            OneNoteXml.ValidatePostInsertCaret(update);

            XDocument modified = new XDocument(page);
            OneNoteXml.InsertAtSelection(modified, Fragment());
            modified.Root.Element(one + "QuickStyleDef").SetAttributeValue("fontSize", "99");
            Throws<InvalidOperationException>(delegate
            {
                InsertSnippetOperation.CreateScopedUpdate(
                    page, modified, page.Root.Element(one + "Outline"));
            });
        }

        private static void UnidentifiedTargetNeverUpdates()
        {
            foreach (string invalid in new[] { "page-id", "outline-id", "duplicate" })
            {
                var calls = new List<string>();
                var application = new FakeApplication(calls);
                XDocument page = Page();
                if (invalid == "page-id") page.Root.Attribute("ID").Remove();
                if (invalid == "outline-id")
                    page.Root.Element(OneNoteXml.One + "Outline").Attribute("objectID").Remove();
                if (invalid == "duplicate")
                    page.Root.Add(new XElement(OneNoteXml.One + "Outline",
                        new XAttribute("objectID", "target-outline")));
                using (var client = CreateClient(application, new FakeWindowActivator(calls)))
                {
                    Throws<InvalidOperationException>(delegate
                    {
                        InsertSnippetOperation.Execute(
                            client, page, new Snippet { Id = "text" },
                            delegate { return Fragment(); });
                    });
                }
                Equal(0, calls.Count);
            }
        }

        private static void ScopedBlockInsertRetainsCompleteTargetContent()
        {
            XNamespace one = OneNoteXml.One;
            var calls = new List<string>();
            var application = new FakeApplication(calls);
            XDocument page = Page();
            XElement outline = page.Root.Element(one + "Outline");
            outline.AddFirst(new XElement(one + "Position",
                new XAttribute("x", "100"), new XAttribute("y", "200")),
                new XElement(one + "Size", new XAttribute("width", "300"),
                    new XAttribute("height", "100")));
            outline.Element(one + "OEChildren").Add(
                new XElement(one + "OE", new XAttribute("objectID", "existing-paragraph"),
                    new XElement(one + "T", "Keep this"),
                    new XElement(one + "Image", new XAttribute("format", "png"),
                        new XElement(one + "Data", "AQID"))));
            XDocument fragment = Fragment();
            fragment.Root.Element(OneNoteXml.Snippet + "Content").Add(
                new XElement(one + "OE", new XElement(one + "T", "Second paragraph")));
            string original = page.ToString();
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                InsertSnippetOperation.Execute(
                    client, page, new Snippet { Id = "blocks" }, delegate { return fragment; });
            }
            Equal(original, page.ToString());
            XDocument update = XDocument.Parse(application.UpdatedXml);
            Equal("100", (string)update.Root.Element(one + "Outline")
                .Element(one + "Position").Attribute("x"));
            Equal("snippet,Second paragraph,,Keep this",
                string.Join(",", update.Descendants(one + "T").Select(text => text.Value)));
            Equal("AQID", update.Descendants(one + "Image").Single().Element(one + "Data").Value);
            Equal("existing-paragraph", (string)update.Descendants(one + "Image")
                .Single().Parent.Attribute("objectID"));
            OneNoteXml.ValidatePostInsertCaret(update);
        }

        private static void FailedScopedUpdateLeavesSourceUnchanged()
        {
            var calls = new List<string>();
            var application = new FakeApplication(calls)
            {
                UpdateException = new COMException("Concurrent modification")
            };
            XDocument page = Page();
            string original = page.ToString();
            using (var client = CreateClient(application, new FakeWindowActivator(calls)))
            {
                Throws<COMException>(delegate
                {
                    InsertSnippetOperation.Execute(
                        client, page, new Snippet { Id = "text" }, delegate { return Fragment(); });
                });
            }
            Equal(original, page.ToString());
            Equal("update", string.Join(",", calls));
        }

        private static void ExpirationSurfaceIsRemoved()
        {
            var assembly = typeof(Schnellbausteine.AddIn.AddIn).Assembly;
            Equal(null, assembly.GetType("Schnellbausteine.AddIn.ExpirationPolicy", false));
            Equal(null, assembly.GetType("Schnellbausteine.AddIn.BetaNoticeForm", false));
            Equal(null, typeof(Schnellbausteine.AddIn.AddIn).GetMethod("ShowBetaDownload"));
            Equal(null, typeof(Schnellbausteine.AddIn.AddIn).GetMethod("GetFeatureEnabled"));
        }

        private static void FeatureActionsRemainAvailableWithoutClock()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                Snippet snippet = repository.Save("Unbefristet", Fragment(), false);
                var addIn = new Schnellbausteine.AddIn.AddIn(repository);
                var inserted = new AutoResetEvent(false);
                SetAction(addIn, "insertSnippetAction", new Action<Snippet, IntPtr>(
                    delegate(Snippet value, IntPtr ownerHandle)
                    {
                        Equal(snippet.Id, value.Id);
                        inserted.Set();
                    }));
                try
                {
                    XDocument menu = XDocument.Parse(addIn.GetSnippetMenuContent(
                        new FakeRibbonControl("sqMain")));
                    string controlId = (string)menu.Root.Elements()
                        .First(element => (string)element.Attribute("onAction") == "InsertSnippet")
                        .Attribute("id");
                    addIn.InsertSnippet(new FakeRibbonControl(controlId));
                    Equal(true, inserted.WaitOne(TimeSpan.FromSeconds(5)));
                    string ribbon = addIn.GetCustomUI(null);
                    Equal(false, ribbon.Contains("GetFeatureEnabled"));
                    Equal(false, ribbon.Contains("ShowBetaDownload"));
                }
                finally
                {
                    Disconnect(addIn);
                    inserted.Dispose();
                }
            });
        }
        private static void AssertManageMenus(Schnellbausteine.AddIn.AddIn addIn)
        {
            string[] menuIds =
            {
                "sqMain",
                "sqInsertText",
                "sqInsertOutline",
                "sqInsertInk",
                "sqInsertLasso"
            };
            foreach (string menuId in menuIds)
            {
                string contentXml = addIn.GetSnippetMenuContent(
                    new FakeRibbonControl(menuId));
                XElement[] items = XDocument.Parse(contentXml)
                    .Root.Elements().ToArray();
                Equal("menuSeparator", items[items.Length - 2].Name.LocalName);
                Equal("Schnellbausteine verwalten",
                    (string)items[items.Length - 1].Attribute("label"));
                Equal("ManageSnippets",
                    (string)items[items.Length - 1].Attribute("onAction"));
                Equal(false,
                    contentXml.Contains("OneNote Schnellbausteine verwalten"));
            }
        }

        private static void SetAction(object instance, string fieldName, object action)
        {
            SetField(instance, fieldName, action);
        }

        private static void SetField(object instance, string fieldName, object value)
        {
            instance.GetType().GetField(
                fieldName,
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic).SetValue(instance, value);
        }

        private static void Disconnect(Schnellbausteine.AddIn.AddIn addIn)
        {
            Array custom = new object[0];
            addIn.OnDisconnection(
                Extensibility.ext_DisconnectMode.ext_dm_HostShutdown,
                ref custom);
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
                "<one:Page xmlns:one=\"" + OneNoteXml.One + "\" ID=\"page-id\">" +
                "<one:Outline objectID=\"target-outline\"><one:OEChildren><one:OE>" +
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
