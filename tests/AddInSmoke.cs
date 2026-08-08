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
            try
            {
                Assembly assembly = Assembly.LoadFrom(args[0]);
                Type type = assembly.GetType("Schnellbausteine.AddIn.AddIn", true);
                Type clientType = assembly.GetType(
                    "Schnellbausteine.AddIn.OneNoteClient", true);
                FieldInfo applicationField = clientType.GetField(
                    "application", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo dispatcherField = type.GetField(
                    "dispatcher", BindingFlags.Instance | BindingFlags.NonPublic);
                Type dispatcherType = assembly.GetType(
                    "Schnellbausteine.AddIn.StaDispatcher", true);
                MethodInfo dispatchMethod = type.GetMethod(
                    "DispatchUserAction", BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo selectionMethod = type.GetMethod(
                    "HasCurrentSelection", BindingFlags.Instance | BindingFlags.NonPublic);
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
                bool callbackHarnessWorks = VerifyCallbackDispatch(addIn, dispatchMethod);
                bool visibilityCallback = (bool)type.GetMethod("CanSaveSelection").Invoke(
                    addIn, new object[] { null });
                string ribbonXml = (string)type.GetMethod("GetCustomUI").Invoke(
                    addIn, new object[] { "Microsoft.OneNote" });
                string menuXml = (string)type.GetMethod("GetSnippetMenuContent").Invoke(
                    addIn, new object[] { null });

                XNamespace ribbonNamespace =
                    "http://schemas.microsoft.com/office/2009/07/customui";
                XDocument ribbon = XDocument.Parse(ribbonXml);
                XElement contextMenus = ribbon.Root.Element(ribbonNamespace + "contextMenus");
                XElement textMenu = contextMenus == null
                    ? null
                    : contextMenus.Elements(ribbonNamespace + "contextMenu").SingleOrDefault(
                        element => (string)element.Attribute("idMso") == "ContextMenuText");
                XElement saveButton = textMenu == null
                    ? null
                    : textMenu.Elements(ribbonNamespace + "button").SingleOrDefault(
                        element => (string)element.Attribute("id") == "sqSaveSelection");

                if (!ribbonXml.Contains("sqInsertMenu") ||
                    !ribbonXml.Contains("label=\"Einfügen\"") ||
                    saveButton == null ||
                    (string)saveButton.Attribute("label") != "Als Schnellbaustein speichern..." ||
                    (string)saveButton.Attribute("imageMso") !=
                        "SaveSelectionToQuickPartGallery" ||
                    (string)saveButton.Attribute("onAction") != "SaveSelection" ||
                    (string)saveButton.Attribute("insertBeforeMso") != "Cut" ||
                    saveButton.Attribute("getVisible") != null ||
                    applicationField == null ||
                    applicationField.FieldType.FullName !=
                        "Microsoft.Office.Interop.OneNote.IApplication" ||
                    dispatcherField == null ||
                    dispatcherField.FieldType != dispatcherType ||
                    dispatcherType.GetMethod("Post") == null ||
                    dispatcherType.GetMethod("PostAfterCallback") == null ||
                    dispatcherType.GetMethod("Shutdown") == null ||
                    !dispatcherWorks ||
                    !callbackHarnessWorks ||
                    dispatchMethod == null ||
                    !CallsMethod(type.GetMethod("SaveSelection"), dispatchMethod) ||
                    !CallsMethod(type.GetMethod("InsertSnippet"), dispatchMethod) ||
                    !CallsMethod(type.GetMethod("ManageSnippets"), dispatchMethod) ||
                    hasDynamicCallSites ||
                    referencesDynamicBinder ||
                    !visibilityCallback ||
                    type.GetMethod("CanSaveSelection") == null ||
                    selectionMethod == null ||
                    type.GetMethod("SaveSelection") == null ||
                    !menuXml.Contains("sqEmpty"))
                {
                    Console.Error.WriteLine("Ribbon content is incomplete.");
                    return 1;
                }

                if (!VerifyDiagnostics(diagnosticLog))
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
                Environment.SetEnvironmentVariable("SCHNELLBAUSTEINE_LOG_PATH", null);
                if (File.Exists(diagnosticLog)) File.Delete(diagnosticLog);
            }
        }

        private static bool CallsMethod(MethodInfo caller, MethodInfo called)
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

        private static bool VerifyCallbackDispatch(object addIn, MethodInfo dispatchMethod)
        {
            var callbackBodyCompleted = new TaskCompletionSource<object>();
            var actionCompleted = new ManualResetEvent(false);
            ApartmentState apartment = ApartmentState.Unknown;
            int callingThreadId = Thread.CurrentThread.ManagedThreadId;
            int dispatchedThreadId = callingThreadId;
            bool result = false;
            bool shutdown = false;
            try
            {
                dispatchMethod.Invoke(addIn, new object[]
                {
                    "Harness",
                    callbackBodyCompleted.Task,
                    "Harness error",
                    new Action<IntPtr>(delegate
                    {
                        apartment = Thread.CurrentThread.GetApartmentState();
                        dispatchedThreadId = Thread.CurrentThread.ManagedThreadId;
                        actionCompleted.Set();
                    })
                });
                if (actionCompleted.WaitOne(TimeSpan.FromMilliseconds(150))) return false;
                callbackBodyCompleted.SetResult(null);
                result = actionCompleted.WaitOne(TimeSpan.FromSeconds(5)) &&
                    apartment == ApartmentState.STA &&
                    dispatchedThreadId != callingThreadId;
            }
            finally
            {
                callbackBodyCompleted.TrySetResult(null);
                actionCompleted.Dispose();
                FieldInfo dispatcherField = addIn.GetType().GetField(
                    "dispatcher", BindingFlags.Instance | BindingFlags.NonPublic);
                object dispatcher = dispatcherField.GetValue(addIn);
                shutdown = ShutdownDispatcher(dispatcher, dispatcher.GetType());
            }
            return result && shutdown;
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
                "operation=dispatch.Harness phase=entry",
                "operation=dispatch.Harness phase=action.entry",
                "operation=dispatch.Harness phase=action.exit"
            };
            string harnessAction = log.Split(new[] { Environment.NewLine },
                StringSplitOptions.RemoveEmptyEntries).SingleOrDefault(
                    line => line.Contains("operation=dispatch.Harness phase=action.entry"));
            return required.All(log.Contains) &&
                harnessAction != null &&
                harnessAction.Contains("apartment=STA");
        }

        private static bool ShutdownDispatcher(object dispatcher, Type dispatcherType)
        {
            return (bool)dispatcherType.GetMethod("Shutdown").Invoke(
                dispatcher, new object[] { TimeSpan.FromSeconds(5) });
        }
    }
}
