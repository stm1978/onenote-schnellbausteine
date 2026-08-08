using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Schnellbausteine.Tests
{
    internal static class AddInSmoke
    {
        private sealed class MtaInvocation
        {
            public bool Completed;
            public TimeSpan Duration;
            public Exception Exception;
            public Thread Thread;
            public int ThreadId;
        }

        private static int Main(string[] args)
        {
            if (args.Length < 1 || args.Length > 2)
            {
                Console.Error.WriteLine(
                    "Usage: Schnellbausteine.AddInSmoke.exe <add-in-path> [--require-selection]");
                return 1;
            }

            string diagnosticLog =
                "Schnellbausteine-Smoke-" + Guid.NewGuid().ToString("N") + ".log";
            Environment.SetEnvironmentVariable("SCHNELLBAUSTEINE_LOG_PATH", diagnosticLog);
            Type logType = null;
            try
            {
                Assembly assembly = Assembly.LoadFrom(args[0]);
                Type type = assembly.GetType("Schnellbausteine.AddIn.AddIn", true);
                Type clientType = assembly.GetType(
                    "Schnellbausteine.AddIn.OneNoteClient", true);
                Type applicationAdapterType = assembly.GetType(
                    "Schnellbausteine.AddIn.OneNoteApplicationAdapter", true);
                Type applicationAdapterInterfaceType = assembly.GetType(
                    "Schnellbausteine.AddIn.IOneNoteApplicationAdapter", true);
                Type nameDialogType = assembly.GetType(
                    "Schnellbausteine.AddIn.NameDialog", true);
                FieldInfo applicationField = clientType.GetField(
                    "application", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo piaApplicationField = applicationAdapterType.GetField(
                    "application", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo dispatcherField = type.GetField(
                    "dispatcher", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo saveSelectionActionField = type.GetField(
                    "saveSelectionAction", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo insertSnippetActionField = type.GetField(
                    "insertSnippetAction", BindingFlags.Instance | BindingFlags.NonPublic);
                Type dispatcherType = assembly.GetType(
                    "Schnellbausteine.AddIn.StaDispatcher", true);
                logType = assembly.GetType("Schnellbausteine.AddIn.Log", true);
                MethodInfo dispatchMethod = type.GetMethod(
                    "DispatchUserAction", BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo selectionMethod = type.GetMethod(
                    "HasCurrentSelection", BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo saveSelectionMethod = type.GetMethod("SaveSelection");
                ConstructorInfo[] forbiddenCallbackConstructors =
                    clientType.GetConstructors(
                        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                    .Concat(nameDialogType.GetConstructors(
                        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                    .ToArray();
                bool hasDynamicCallSites = assembly.GetTypes()
                    .SelectMany(candidate => candidate.GetFields(
                        BindingFlags.Public | BindingFlags.NonPublic |
                        BindingFlags.Static | BindingFlags.Instance))
                    .Any(field => field.FieldType.IsGenericType &&
                        field.FieldType.GetGenericTypeDefinition() == typeof(CallSite<>));
                bool referencesDynamicBinder = assembly.GetReferencedAssemblies()
                    .Any(reference => reference.Name == "Microsoft.CSharp");
                object addIn = Activator.CreateInstance(type);
                bool dispatcherWorks = VerifyDispatcher(
                    Activator.CreateInstance(dispatcherType), dispatcherType);
                bool timeoutGuardWorks = VerifyCallbackTimeoutGuard();
                bool saveCallbackHarnessWorks = VerifyDeferredCallback(
                    addIn, saveSelectionActionField, "SaveSelection");
                object insertAddIn = Activator.CreateInstance(type);
                bool insertCallbackHarnessWorks = VerifyDeferredCallback(
                    insertAddIn, insertSnippetActionField, "InsertSnippet");
                bool visibilityCallback = (bool)type.GetMethod("CanSaveSelection").Invoke(
                    addIn, new object[] { null });
                string ribbonXml = (string)type.GetMethod("GetCustomUI").Invoke(
                    addIn, new object[] { "Microsoft.OneNote" });
                XNamespace ribbonNamespace =
                    "http://schemas.microsoft.com/office/2009/07/customui";
                XDocument ribbon = XDocument.Parse(ribbonXml);
                XElement insertButton = ribbon.Descendants(ribbonNamespace + "button")
                    .SingleOrDefault(element => (string)element.Attribute("id") == "sqInsert");
                XElement contextMenus = ribbon.Root.Element(ribbonNamespace + "contextMenus");
                XElement textMenu = contextMenus == null
                    ? null
                    : contextMenus.Elements(ribbonNamespace + "contextMenu").SingleOrDefault(
                        element => (string)element.Attribute("idMso") == "ContextMenuText");
                XElement saveButton = textMenu == null
                    ? null
                    : textMenu.Elements(ribbonNamespace + "button").SingleOrDefault(
                        element => (string)element.Attribute("id") == "sqSaveSelection");

                if (insertButton == null ||
                    (string)insertButton.Attribute("label") != "Einfügen..." ||
                    (string)insertButton.Attribute("onAction") != "InsertSnippet" ||
                    ribbon.Descendants(ribbonNamespace + "dynamicMenu").Any() ||
                    saveButton == null ||
                    (string)saveButton.Attribute("label") != "Als Schnellbaustein speichern..." ||
                    (string)saveButton.Attribute("imageMso") !=
                        "SaveSelectionToQuickPartGallery" ||
                    (string)saveButton.Attribute("onAction") != "SaveSelection" ||
                    (string)saveButton.Attribute("insertBeforeMso") != "Cut" ||
                    saveButton.Attribute("getVisible") != null ||
                    applicationField == null ||
                    applicationField.FieldType != applicationAdapterInterfaceType ||
                    piaApplicationField == null ||
                    piaApplicationField.FieldType.FullName !=
                        "Microsoft.Office.Interop.OneNote.IApplication" ||
                    dispatcherField == null ||
                    dispatcherField.FieldType != dispatcherType ||
                    saveSelectionActionField == null ||
                    saveSelectionActionField.FieldType != typeof(Action<IntPtr>) ||
                    insertSnippetActionField == null ||
                    insertSnippetActionField.FieldType != typeof(Action<IntPtr>) ||
                    dispatcherType.GetMethod("Post") == null ||
                    dispatcherType.GetMethod("PostAfterCallback") == null ||
                    dispatcherType.GetMethod("Shutdown") == null ||
                    !dispatcherWorks ||
                    !timeoutGuardWorks ||
                    !saveCallbackHarnessWorks ||
                    !insertCallbackHarnessWorks ||
                    dispatchMethod == null ||
                    !CallsMember(saveSelectionMethod, dispatchMethod) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => CallsMember(saveSelectionMethod, constructor)) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => CallsMember(type.GetMethod("InsertSnippet"), constructor)) ||
                    !CallsMethod(type.GetMethod("InsertSnippet"), dispatchMethod) ||
                    !CallsMethod(type.GetMethod("ManageSnippets"), dispatchMethod) ||
                    hasDynamicCallSites ||
                    referencesDynamicBinder ||
                    !visibilityCallback ||
                    type.GetMethod("CanSaveSelection") == null ||
                    selectionMethod == null ||
                    type.GetMethod("SaveSelection") == null ||
                    type.GetMethod("GetSnippetMenuContent") != null)
                {
                    Console.Error.WriteLine("Ribbon content is incomplete.");
                    return 1;
                }

                bool logFlushed = (bool)logType.GetMethod("Flush").Invoke(
                    null, new object[] { TimeSpan.FromSeconds(5) });
                if (!logFlushed || !VerifyDiagnostics(diagnosticLog))
                {
                    Console.Error.WriteLine("Callback diagnostic logging is incomplete.");
                    return 1;
                }

                if (args.Length == 2)
                {
                    if (args[1] != "--require-selection")
                    {
                        Console.Error.WriteLine("Unknown option: " + args[1]);
                        return 1;
                    }

                    bool hasSelection = (bool)selectionMethod.Invoke(addIn, new object[0]);
                    if (!hasSelection)
                    {
                        Console.Error.WriteLine(
                            "OneNote did not return a readable active text selection.");
                        return 1;
                    }
                }

                Console.WriteLine("Managed add-in and Ribbon smoke test succeeded.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
            finally
            {
                if (logType != null)
                {
                    logType.GetMethod("Flush").Invoke(
                        null, new object[] { TimeSpan.FromSeconds(5) });
                }
                Environment.SetEnvironmentVariable("SCHNELLBAUSTEINE_LOG_PATH", null);
                if (File.Exists(diagnosticLog)) File.Delete(diagnosticLog);
            }
        }

        private static bool CallsMethod(MethodInfo caller, MethodInfo called)
        {
            return CallsMember(caller, called);
        }

        private static bool CallsMember(MethodInfo caller, MethodBase called)
        {
            byte[] body = caller.GetMethodBody().GetILAsByteArray();
            byte[] token = BitConverter.GetBytes(called.MetadataToken);
            for (int index = 0; index <= body.Length - token.Length; index++)
            {
                bool match = true;
                for (int offset = 0; offset < token.Length; offset++)
                {
                    if (body[index + offset] != token[offset])
                    {
                        match = false;
                        break;
                    }
                }
                if (match) return true;
            }
            return false;
        }

        private static bool VerifyDispatcher(object dispatcher, Type dispatcherType)
        {
            var completed = new ManualResetEvent(false);
            ApartmentState apartment = ApartmentState.Unknown;
            int callingThreadId = Thread.CurrentThread.ManagedThreadId;
            int dispatchedThreadId = callingThreadId;
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            var stopwatch = new Stopwatch();
            bool result = false;
            bool shutdown = false;
            try
            {
                dispatcherType.GetMethod("PostAfterCallback").Invoke(dispatcher, new object[]
                {
                    callbackBodyCompleted.Task,
                    new Action(delegate
                    {
                        apartment = Thread.CurrentThread.GetApartmentState();
                        dispatchedThreadId = Thread.CurrentThread.ManagedThreadId;
                        completed.Set();
                    }),
                    new Action<Exception>(delegate { completed.Set(); })
                });
                if (completed.WaitOne(TimeSpan.FromMilliseconds(150))) return false;
                stopwatch.Start();
                callbackBodyCompleted.SetResult(null);
                result = completed.WaitOne(TimeSpan.FromSeconds(5)) &&
                    stopwatch.Elapsed >= TimeSpan.FromMilliseconds(75) &&
                    apartment == ApartmentState.STA &&
                    dispatchedThreadId != callingThreadId;
            }
            finally
            {
                callbackBodyCompleted.TrySetResult(null);
                completed.Dispose();
                shutdown = ShutdownDispatcher(dispatcher, dispatcherType);
            }
            return result && shutdown;
        }

        private static bool VerifyDeferredCallback(
            object addIn, FieldInfo actionField, string callbackName)
        {
            var actionCompleted = new ManualResetEvent(false);
            ApartmentState apartment = ApartmentState.Unknown;
            int callbackThreadId = 0;
            int dispatchedThreadId = 0;
            bool callbackReturned = false;
            bool actionStartedBeforeReturn = false;
            bool result = false;
            bool shutdown = false;
            bool callbackFinished = false;
            try
            {
                actionField.SetValue(addIn, new Action<IntPtr>(delegate
                {
                    actionStartedBeforeReturn = !Volatile.Read(ref callbackReturned);
                    apartment = Thread.CurrentThread.GetApartmentState();
                    dispatchedThreadId = Thread.CurrentThread.ManagedThreadId;
                    actionCompleted.Set();
                }));
                MtaInvocation invocation = InvokeOnMta(delegate
                {
                    try
                    {
                        addIn.GetType().GetMethod(callbackName).Invoke(
                            addIn, new object[] { null });
                    }
                    finally
                    {
                        Volatile.Write(ref callbackReturned, true);
                    }
                }, TimeSpan.FromSeconds(2));
                callbackThreadId = invocation.ThreadId;
                if (!invocation.Completed) return false;
                callbackFinished = true;
                result = actionCompleted.WaitOne(TimeSpan.FromSeconds(5)) &&
                    invocation.Exception == null &&
                    !actionStartedBeforeReturn &&
                    invocation.Duration < TimeSpan.FromSeconds(2) &&
                    apartment == ApartmentState.STA &&
                    dispatchedThreadId != callbackThreadId;
            }
            finally
            {
                if (callbackFinished) actionCompleted.Dispose();
                FieldInfo dispatcherField = addIn.GetType().GetField(
                    "dispatcher", BindingFlags.Instance | BindingFlags.NonPublic);
                object dispatcher = dispatcherField.GetValue(addIn);
                shutdown = ShutdownDispatcher(dispatcher, dispatcher.GetType());
            }
            return result && shutdown;
        }

        private static bool VerifyCallbackTimeoutGuard()
        {
            var release = new ManualResetEvent(false);
            MtaInvocation invocation = InvokeOnMta(delegate
            {
                release.WaitOne();
            }, TimeSpan.FromMilliseconds(150));
            release.Set();
            bool cleanedUp = invocation.Thread.Join(TimeSpan.FromSeconds(2));
            release.Dispose();
            return !invocation.Completed && cleanedUp;
        }

        private static bool VerifyDiagnostics(string path)
        {
            string log = File.ReadAllText(path);
            string[] required =
            {
                "pid=",
                "thread=",
                "apartment=",
                "operation=dispatcher phase=thread.ready",
                "operation=dispatcher phase=deferred.wait-callback",
                "operation=dispatcher phase=deferred.callback-complete",
                "operation=dispatcher phase=action.before",
                "operation=dispatcher phase=action.after",
                "operation=dispatcher phase=thread.exit",
                "operation=dispatcher phase=shutdown.complete",
                "operation=callback.SaveSelection phase=entry",
                "operation=callback.SaveSelection phase=exit",
                "operation=dispatch.SaveSelection phase=entry",
                "operation=dispatch.SaveSelection phase=action.entry",
                "operation=dispatch.SaveSelection phase=action.exit"
            };
            string harnessAction = log.Split(new[] { Environment.NewLine },
                StringSplitOptions.RemoveEmptyEntries).SingleOrDefault(
                    line => line.Contains("operation=dispatch.SaveSelection phase=action.entry"));
            return required.All(log.Contains) &&
                harnessAction != null &&
                harnessAction.Contains("apartment=STA");
        }

        private static bool ShutdownDispatcher(object dispatcher, Type dispatcherType)
        {
            return (bool)dispatcherType.GetMethod("Shutdown").Invoke(
                dispatcher, new object[] { TimeSpan.FromSeconds(5) });
        }

        private static MtaInvocation InvokeOnMta(Action callback, TimeSpan timeout)
        {
            var invocation = new MtaInvocation();
            var stopwatch = new Stopwatch();
            var thread = new Thread(new ThreadStart(delegate
            {
                invocation.ThreadId = Thread.CurrentThread.ManagedThreadId;
                try
                {
                    stopwatch.Start();
                    callback();
                }
                catch (Exception exception)
                {
                    invocation.Exception = exception;
                }
                finally
                {
                    stopwatch.Stop();
                }
            }));
            invocation.Thread = thread;
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.MTA);
            thread.Start();
            invocation.Completed = thread.Join(timeout);
            invocation.Duration = stopwatch.Elapsed;
            return invocation;
        }
    }
}
