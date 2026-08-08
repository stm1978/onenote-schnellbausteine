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
        private sealed class FakeClock : ILocalClock
        {
            public DateTimeOffset Value;

            public DateTimeOffset LocalNow
            {
                get { return Value; }
            }
        }

        private sealed class FakeNoticePresenter : IBetaNoticePresenter
        {
            public readonly AutoResetEvent Shown = new AutoResetEvent(false);
            public int Count;
            public BetaStatus Status;

            public void Show(BetaStatus status, IntPtr ownerHandle)
            {
                Status = status;
                Interlocked.Increment(ref Count);
                Shown.Set();
            }
        }

        private sealed class FakeWebsiteLauncher : IWebsiteLauncher
        {
            public Uri Website;
            public Exception Exception;

            public void Open(Uri website)
            {
                Website = website;
                if (Exception != null) throw Exception;
            }
        }

        private sealed class FakeApplication : IOneNoteApplicationAdapter
        {
            public readonly List<string> Calls;
            public Exception UpdateException;
            public string BinaryData;

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
                Calls.Add("binary:" + pageId + ":" + callbackId);
                return BinaryData;
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
                MenuItemsAreSortedEscapedAndAddressable();
                EmptyMenuItemIsDisabled();
                ActivationFailureIsNonModalAndDoesNotFailInsert();
                MenuInvalidationRejectsStaleIdsAndReloadsCrud();
                TargetAlreadyForegroundAcceptsSetForegroundFalse();
                CurrentForegroundAndTargetThreadsAreAttached();
                AttachedThreadsAreDetachedInFinally();
                MinimizedTargetIsRestoredBeforeActivation();
                RejectedActivationDoesNotFailInsert();
                BinaryCallbacksUseCapturedPageOnClientThread();
                ExpirationPolicyUsesExactLocalBoundaries();
                StartupNoticeIsDeferredOncePerLifecycle();
                StartupNoticeIsAbsentBeforeWarningPeriod();
                ExpiredStartupNoticeIsDeferred();
                ShutdownCancelsQueuedStartupNotice();
                ExpiredActionsAreBlockedBeforeCommandBodies();
                WarningActionsRemainEnabled();
                ExpiredMenuHidesNamesAndStaleItemsAreBlocked();
                ExpiredDownloadCallbackShowsNotice();
                WebsiteUsesExactHttpsUriAndSurfacesErrors();
                Console.WriteLine("23/23 OneNote-Client-Harness-Tests erfolgreich.");
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
                Equal("OneNote Schnellbausteine verwalten",
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
                Equal("OneNote Schnellbausteine verwalten",
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

        private static void ExpirationPolicyUsesExactLocalBoundaries()
        {
            var clock = new FakeClock();
            var policy = new ExpirationPolicy(clock);

            clock.Value = new DateTimeOffset(
                2026, 11, 30, 23, 59, 59, 999, TimeSpan.FromHours(14));
            Equal(BetaStatus.Active, policy.CurrentStatus);
            Equal(true, policy.AllowsFeatures);

            clock.Value = new DateTimeOffset(
                2026, 12, 1, 0, 0, 0, TimeSpan.FromHours(-8));
            Equal(BetaStatus.Warning, policy.CurrentStatus);
            Equal(true, policy.AllowsFeatures);

            clock.Value = new DateTimeOffset(
                2026, 12, 31, 23, 59, 59, 999, TimeSpan.FromHours(14));
            Equal(BetaStatus.Warning, policy.CurrentStatus);
            Equal(true, policy.AllowsFeatures);

            clock.Value = new DateTimeOffset(
                2027, 1, 1, 0, 0, 0, TimeSpan.FromHours(-8));
            Equal(BetaStatus.Expired, policy.CurrentStatus);
            Equal(false, policy.AllowsFeatures);
        }

        private static void StartupNoticeIsDeferredOncePerLifecycle()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(
                        2026, 12, 1, 0, 0, 0, TimeSpan.FromHours(1))
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                try
                {
                    Array custom = new object[0];
                    addIn.OnStartupComplete(ref custom);
                    Equal(0, presenter.Count);
                    Equal(true, presenter.Shown.WaitOne(TimeSpan.FromSeconds(5)));
                    Equal(BetaStatus.Warning, presenter.Status);

                    addIn.OnStartupComplete(ref custom);
                    Equal(false, presenter.Shown.WaitOne(TimeSpan.FromMilliseconds(300)));
                    Equal(1, presenter.Count);
                }
                finally
                {
                    Disconnect(addIn);
                    presenter.Shown.Dispose();
                }
            });
        }

        private static void StartupNoticeIsAbsentBeforeWarningPeriod()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(
                        2026, 11, 30, 23, 59, 59, 999, TimeSpan.Zero)
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                try
                {
                    Array custom = new object[0];
                    addIn.OnStartupComplete(ref custom);
                    Equal(false, presenter.Shown.WaitOne(TimeSpan.FromMilliseconds(300)));
                    Equal(0, presenter.Count);
                }
                finally
                {
                    Disconnect(addIn);
                    presenter.Shown.Dispose();
                }
            });
        }

        private static void ExpiredActionsAreBlockedBeforeCommandBodies()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(2027, 1, 1, 0, 0, 0, TimeSpan.Zero)
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                int commandCalls = 0;
                SetAction(addIn, "saveSelectionAction", new Action<IntPtr>(
                    delegate { Interlocked.Increment(ref commandCalls); }));
                SetAction(addIn, "manageSnippetsAction", new Action<IntPtr>(
                    delegate { Interlocked.Increment(ref commandCalls); }));
                try
                {
                    Equal(false, addIn.GetFeatureEnabled(null));
                    addIn.SaveSelection(null);
                    Equal(true, presenter.Shown.WaitOne(TimeSpan.FromSeconds(5)));
                    addIn.ManageSnippets(null);
                    Equal(true, presenter.Shown.WaitOne(TimeSpan.FromSeconds(5)));
                    Equal(0, commandCalls);
                    Equal(2, presenter.Count);
                }
                finally
                {
                    Disconnect(addIn);
                    presenter.Shown.Dispose();
                }
            });
        }

        private static void ExpiredStartupNoticeIsDeferred()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(2027, 1, 1, 0, 0, 0, TimeSpan.Zero)
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                try
                {
                    Array custom = new object[0];
                    addIn.OnStartupComplete(ref custom);
                    Equal(0, presenter.Count);
                    Equal(true, presenter.Shown.WaitOne(TimeSpan.FromSeconds(5)));
                    Equal(BetaStatus.Expired, presenter.Status);
                }
                finally
                {
                    Disconnect(addIn);
                    presenter.Shown.Dispose();
                }
            });
        }

        private static void ShutdownCancelsQueuedStartupNotice()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(2026, 12, 1, 0, 0, 0, TimeSpan.Zero)
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                try
                {
                    Array custom = new object[0];
                    addIn.OnStartupComplete(ref custom);
                    addIn.OnBeginShutdown(ref custom);
                    Equal(false, presenter.Shown.WaitOne(TimeSpan.FromMilliseconds(400)));
                    Equal(0, presenter.Count);
                }
                finally
                {
                    Disconnect(addIn);
                    presenter.Shown.Dispose();
                }
            });
        }

        private static void WarningActionsRemainEnabled()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(2026, 12, 31, 23, 59, 59, 999, TimeSpan.Zero)
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                var called = new AutoResetEvent(false);
                SetAction(addIn, "saveSelectionAction", new Action<IntPtr>(
                    delegate { called.Set(); }));
                try
                {
                    Equal(true, addIn.GetFeatureEnabled(null));
                    addIn.SaveSelection(null);
                    Equal(true, called.WaitOne(TimeSpan.FromSeconds(5)));
                    Equal(0, presenter.Count);
                }
                finally
                {
                    Disconnect(addIn);
                    called.Dispose();
                    presenter.Shown.Dispose();
                }
            });
        }

        private static void ExpiredMenuHidesNamesAndStaleItemsAreBlocked()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                repository.Save("Vertraulicher Name", Fragment(), false);
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(2026, 12, 31, 23, 0, 0, TimeSpan.Zero)
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                var menu = new SnippetMenuModel(repository);
                XDocument active = XDocument.Parse(
                    menu.RefreshAndGetContent("sqMain"));
                string staleId = (string)active.Root.Elements().First().Attribute("id");
                SetField(addIn, "menu", menu);
                clock.Value = new DateTimeOffset(2027, 1, 1, 0, 0, 0, TimeSpan.Zero);
                int insertCalls = 0;
                SetAction(
                    addIn,
                    "insertSnippetAction",
                    new Action<Snippet, IntPtr>(
                        delegate { Interlocked.Increment(ref insertCalls); }));
                try
                {
                    string expiredXml = addIn.GetSnippetMenuContent(null);
                    XDocument expired = XDocument.Parse(expiredXml);
                    XElement[] items = expired.Root.Elements().ToArray();
                    Equal(1, items.Length);
                    Equal("Aktuelle Version herunterladen...",
                        (string)items[0].Attribute("label"));
                    Equal("ShowBetaDownload", (string)items[0].Attribute("onAction"));
                    Equal(false, expiredXml.Contains("Vertraulicher Name"));

                    addIn.InsertSnippet(null);
                    Equal(true, presenter.Shown.WaitOne(TimeSpan.FromSeconds(5)));
                    Equal(0, insertCalls);
                    Throws<InvalidOperationException>(
                        delegate { menu.GetSnippet(staleId); });
                }
                finally
                {
                    Disconnect(addIn);
                    presenter.Shown.Dispose();
                }
            });
        }

        private static void WebsiteUsesExactHttpsUriAndSurfacesErrors()
        {
            var launcher = new FakeWebsiteLauncher();
            Exception visibleError = null;
            BetaWebsite.Open(launcher, delegate(Exception exception)
            {
                visibleError = exception;
            });
            Equal("https://www.malter365.de/", launcher.Website.AbsoluteUri);
            Equal(null, visibleError);

            launcher.Exception = new InvalidOperationException("browser failed");
            BetaWebsite.Open(launcher, delegate(Exception exception)
            {
                visibleError = exception;
            });
            Equal(launcher.Exception, visibleError);
        }

        private static void ExpiredDownloadCallbackShowsNotice()
        {
            WithTemporaryRepository(delegate(SnippetRepository repository)
            {
                var clock = new FakeClock
                {
                    Value = new DateTimeOffset(2027, 1, 1, 0, 0, 0, TimeSpan.Zero)
                };
                var presenter = new FakeNoticePresenter();
                var addIn = new Schnellbausteine.AddIn.AddIn(
                    repository,
                    new ExpirationPolicy(clock),
                    presenter);
                try
                {
                    addIn.ShowBetaDownload(null);
                    Equal(true, presenter.Shown.WaitOne(TimeSpan.FromSeconds(5)));
                    Equal(BetaStatus.Expired, presenter.Status);
                    Equal(1, presenter.Count);
                }
                finally
                {
                    Disconnect(addIn);
                    presenter.Shown.Dispose();
                }
            });
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
