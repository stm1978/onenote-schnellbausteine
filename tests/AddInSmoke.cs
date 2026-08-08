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
            string galleryRepositoryRoot = Path.Combine(
                Path.GetTempPath(),
                "Schnellbausteine.AddInSmoke",
                Guid.NewGuid().ToString("N"));
            Environment.SetEnvironmentVariable("SCHNELLBAUSTEINE_LOG_PATH", diagnosticLog);
            Type logType = null;
            try
            {
                Assembly assembly = Assembly.LoadFrom(args[0]);
                Type type = assembly.GetType("Schnellbausteine.AddIn.AddIn", true);
                Assembly coreAssembly = Assembly.LoadFrom(Path.Combine(
                    Path.GetDirectoryName(Path.GetFullPath(args[0])),
                    "Schnellbausteine.Core.dll"));
                Type repositoryType = coreAssembly.GetType(
                    "Schnellbausteine.Core.SnippetRepository", true);
                Type clientType = assembly.GetType(
                    "Schnellbausteine.AddIn.OneNoteClient", true);
                Type applicationAdapterType = assembly.GetType(
                    "Schnellbausteine.AddIn.OneNoteApplicationAdapter", true);
                Type applicationAdapterInterfaceType = assembly.GetType(
                    "Schnellbausteine.AddIn.IOneNoteApplicationAdapter", true);
                Type nameDialogType = assembly.GetType(
                    "Schnellbausteine.AddIn.NameDialog", true);
                Type manageDialogType = assembly.GetType(
                    "Schnellbausteine.AddIn.ManageDialog", true);
                Type popupType = assembly.GetType(
                    "Schnellbausteine.AddIn.SnippetMenuPopup", false);
                Type popupInterfaceType = assembly.GetType(
                    "Schnellbausteine.AddIn.ISnippetMenuPopup", false);
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
                FieldInfo manageSnippetsActionField = type.GetField(
                    "manageSnippetsAction", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo ribbonUiField = type.GetField(
                    "ribbonUi", BindingFlags.Instance | BindingFlags.NonPublic);
                Type dispatcherType = assembly.GetType(
                    "Schnellbausteine.AddIn.StaDispatcher", true);
                logType = assembly.GetType("Schnellbausteine.AddIn.Log", true);
                MethodInfo dispatchMethod = type.GetMethod(
                    "DispatchUserAction", BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo selectionMethod = type.GetMethod(
                    "HasCurrentSelection", BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo saveSelectionMethod = type.GetMethod("SaveSelection");
                MethodInfo insertSnippetMethod = type.GetMethod("InsertGallerySnippet");
                MethodInfo manageSnippetsMethod = type.GetMethod("ManageSnippets");
                MethodInfo insertSnippetCoreMethod = type.GetMethod(
                    "InsertSnippetCore",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo manageSnippetsCoreMethod = type.GetMethod(
                    "ManageSnippetsCore",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo invalidateGalleryMethod = type.GetMethod(
                    "InvalidateSnippetGallery",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo readCurrentPageMethod = clientType.GetMethod("ReadCurrentPage");
                MethodInfo[] galleryGetterMethods = new[]
                {
                    type.GetMethod("GetSnippetCount"),
                    type.GetMethod("GetSnippetLabel"),
                    type.GetMethod("GetSnippetId"),
                    type.GetMethod("GetSnippetGalleryEnabled")
                };
                ConstructorInfo[] forbiddenCallbackConstructors =
                    clientType.GetConstructors(
                        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                    .Concat(nameDialogType.GetConstructors(
                        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                    .Concat(manageDialogType.GetConstructors(
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
                object manageAddIn = Activator.CreateInstance(type);
                bool manageCallbackHarnessWorks = VerifyDeferredCallback(
                    manageAddIn, manageSnippetsActionField, "ManageSnippets");
                Directory.CreateDirectory(galleryRepositoryRoot);
                object galleryRepository = Activator.CreateInstance(
                    repositoryType,
                    new object[] { galleryRepositoryRoot });
                XDocument galleryFragment = XDocument.Parse(
                    "<sq:Snippet xmlns:sq=\"urn:schnellbausteine:v1\" " +
                    "xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" " +
                    "version=\"1\"><sq:TagDefs/><sq:QuickStyleDefs/><sq:Content>" +
                    "<one:OE><one:T><![CDATA[Test]]></one:T></one:OE>" +
                    "</sq:Content></sq:Snippet>");
                repositoryType.GetMethod("Save").Invoke(
                    galleryRepository,
                    new object[] { "Test", galleryFragment, false });
                ConstructorInfo repositoryConstructor = type.GetConstructor(
                    BindingFlags.Instance | BindingFlags.NonPublic,
                    null,
                    new[] { repositoryType },
                    null);
                object galleryAddIn = repositoryConstructor.Invoke(
                    new[] { galleryRepository });
                bool itemCallbackHarnessWorks = VerifyDeferredGalleryCallback(
                    galleryAddIn,
                    insertSnippetActionField);
                bool visibilityCallback = (bool)type.GetMethod("CanSaveSelection").Invoke(
                    addIn, new object[] { null });
                string ribbonXml = (string)type.GetMethod("GetCustomUI").Invoke(
                    addIn, new object[] { "Microsoft.OneNote" });
                XNamespace ribbonNamespace =
                    "http://schemas.microsoft.com/office/2009/07/customui";
                XDocument ribbon = XDocument.Parse(ribbonXml);
                XElement ribbonGroup = ribbon.Descendants(ribbonNamespace + "group")
                    .SingleOrDefault(element => (string)element.Attribute("id") == "sqGroup");
                XElement mainMenu = ribbonGroup == null
                    ? null
                    : ribbonGroup.Elements(ribbonNamespace + "menu").SingleOrDefault();
                XElement gallery = mainMenu == null
                    ? null
                    : mainMenu.Elements(ribbonNamespace + "gallery").SingleOrDefault();
                XElement separator = mainMenu == null
                    ? null
                    : mainMenu.Elements(ribbonNamespace + "menuSeparator").SingleOrDefault();
                XElement manageButton = mainMenu == null
                    ? null
                    : mainMenu.Elements(ribbonNamespace + "button").SingleOrDefault();
                XElement contextMenus = ribbon.Root.Element(ribbonNamespace + "contextMenus");
                XElement textMenu = contextMenus == null
                    ? null
                    : contextMenus.Elements(ribbonNamespace + "contextMenu").SingleOrDefault(
                        element => (string)element.Attribute("idMso") == "ContextMenuText");
                XElement saveButton = textMenu == null
                    ? null
                    : textMenu.Elements(ribbonNamespace + "button").SingleOrDefault(
                        element => (string)element.Attribute("id") == "sqSaveSelection");

                if (mainMenu == null ||
                    ribbonGroup.Elements().Count() != 1 ||
                    (string)mainMenu.Attribute("id") != "sqMain" ||
                    (string)mainMenu.Attribute("label") != "Schnellbausteine" ||
                    (string)mainMenu.Attribute("size") != "large" ||
                    gallery == null ||
                    (string)gallery.Attribute("id") != "sqGallery" ||
                    (string)gallery.Attribute("columns") != "1" ||
                    (string)gallery.Attribute("itemWidth") != "240" ||
                    (string)gallery.Attribute("itemHeight") != "24" ||
                    (string)gallery.Attribute("showLabel") != "false" ||
                    (string)gallery.Attribute("showItemLabel") != "true" ||
                    (string)gallery.Attribute("showItemImage") != "false" ||
                    (string)gallery.Attribute("getItemCount") != "GetSnippetCount" ||
                    (string)gallery.Attribute("getItemLabel") != "GetSnippetLabel" ||
                    (string)gallery.Attribute("getItemID") != "GetSnippetId" ||
                    (string)gallery.Attribute("getEnabled") != "GetSnippetGalleryEnabled" ||
                    gallery.Attribute("getItemEnabled") != null ||
                    (string)gallery.Attribute("onAction") != "InsertGallerySnippet" ||
                    separator == null ||
                    (string)separator.Attribute("id") != "sqManageSeparator" ||
                    manageButton == null ||
                    (string)manageButton.Attribute("id") != "sqManage" ||
                    (string)manageButton.Attribute("label") != "Schnellbausteine verwalten" ||
                    (string)manageButton.Attribute("onAction") != "ManageSnippets" ||
                    !mainMenu.Elements().SequenceEqual(
                        new[] { gallery, separator, manageButton }) ||
                    ribbon.Descendants(ribbonNamespace + "dynamicMenu").Any() ||
                    ribbon.Descendants(ribbonNamespace + "form").Any() ||
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
                    !insertSnippetActionField.FieldType.IsGenericType ||
                    insertSnippetActionField.FieldType.GetGenericTypeDefinition() != typeof(Action<,>) ||
                    insertSnippetActionField.FieldType.GetGenericArguments()[1] != typeof(IntPtr) ||
                    manageSnippetsActionField == null ||
                    manageSnippetsActionField.FieldType != typeof(Action<IntPtr>) ||
                    ribbonUiField == null ||
                    galleryGetterMethods.Any(method => method == null) ||
                    galleryGetterMethods.Any(method =>
                        method.GetParameters().Length < 1 ||
                        method.GetParameters()[0].ParameterType.FullName !=
                            "Microsoft.Office.Core.IRibbonControl") ||
                    galleryGetterMethods[0].ReturnType != typeof(int) ||
                    galleryGetterMethods[0].GetParameters().Length != 1 ||
                    galleryGetterMethods[1].ReturnType != typeof(string) ||
                    galleryGetterMethods[1].GetParameters().Length != 2 ||
                    galleryGetterMethods[1].GetParameters()[1].ParameterType != typeof(int) ||
                    galleryGetterMethods[2].ReturnType != typeof(string) ||
                    galleryGetterMethods[2].GetParameters().Length != 2 ||
                    galleryGetterMethods[2].GetParameters()[1].ParameterType != typeof(int) ||
                    galleryGetterMethods[3].ReturnType != typeof(bool) ||
                    galleryGetterMethods[3].GetParameters().Length != 1 ||
                    dispatcherType.GetMethod("Post") == null ||
                    dispatcherType.GetMethod("PostAfterCallback") == null ||
                    dispatcherType.GetMethod("Shutdown") == null ||
                    !dispatcherWorks ||
                    !timeoutGuardWorks ||
                    !saveCallbackHarnessWorks ||
                    !manageCallbackHarnessWorks ||
                    !itemCallbackHarnessWorks ||
                    dispatchMethod == null ||
                    !CallsMember(saveSelectionMethod, dispatchMethod) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => CallsMember(saveSelectionMethod, constructor)) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => galleryGetterMethods.Any(
                            getter => CallsMember(getter, constructor))) ||
                    insertSnippetMethod == null ||
                    insertSnippetMethod.ReturnType != typeof(void) ||
                    insertSnippetMethod.GetParameters().Length != 3 ||
                    insertSnippetMethod.GetParameters()[0].ParameterType.FullName !=
                        "Microsoft.Office.Core.IRibbonControl" ||
                    insertSnippetMethod.GetParameters()[1].ParameterType != typeof(string) ||
                    insertSnippetMethod.GetParameters()[2].ParameterType != typeof(int) ||
                    !CallsMethod(insertSnippetMethod, dispatchMethod) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => CallsMember(insertSnippetMethod, constructor)) ||
                    insertSnippetCoreMethod == null ||
                    !CallsMethod(insertSnippetCoreMethod, readCurrentPageMethod) ||
                    manageSnippetsCoreMethod == null ||
                    invalidateGalleryMethod == null ||
                    !CallsMember(saveSelectionMethod.DeclaringType.GetMethod(
                        "SaveSelectionCore",
                        BindingFlags.Instance | BindingFlags.NonPublic),
                        invalidateGalleryMethod) ||
                    !CallsMember(manageSnippetsCoreMethod, invalidateGalleryMethod) ||
                    !CallsMember(
                        invalidateGalleryMethod,
                        ribbonUiField.FieldType.GetMethod("InvalidateControl")) ||
                    popupType != null ||
                    popupInterfaceType != null ||
                    assembly.GetTypes().Any(candidate =>
                        candidate.Name.Contains("SnippetMenuChoice") ||
                        candidate.Name.Contains("SnippetPopupPlacement")) ||
                    assembly.GetTypes().SelectMany(candidate => candidate.GetMethods(
                        BindingFlags.Instance | BindingFlags.Static |
                        BindingFlags.Public | BindingFlags.NonPublic))
                        .Any(method => method.Name == "NavigateTo") ||
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
                if (Directory.Exists(galleryRepositoryRoot))
                {
                    Directory.Delete(galleryRepositoryRoot, true);
                }
            }
        }

        private static bool CallsMethod(MethodInfo caller, MethodInfo called)
        {
            return CallsMember(caller, called);
        }

        private static bool CallsBefore(
            MethodInfo caller,
            MethodBase first,
            MethodBase second)
        {
            return FindMemberCall(caller, first) < FindMemberCall(caller, second) &&
                FindMemberCall(caller, first) >= 0;
        }

        private static int FindMemberCall(MethodInfo caller, MethodBase called)
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
                if (match) return index;
            }
            return -1;
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

        private static bool VerifyDeferredGalleryCallback(
            object addIn,
            FieldInfo actionField)
        {
            var actionCompleted = new ManualResetEvent(false);
            ApartmentState apartment = ApartmentState.Unknown;
            int callbackThreadId = 0;
            int dispatchedThreadId = 0;
            bool callbackReturned = false;
            bool actionStartedBeforeReturn = false;
            bool receivedSnippet = false;
            bool result = false;
            bool callbackFinished = false;
            bool shutdown = false;
            try
            {
                Action<object, IntPtr> action = delegate(object snippet, IntPtr ownerHandle)
                {
                    actionStartedBeforeReturn = !Volatile.Read(ref callbackReturned);
                    apartment = Thread.CurrentThread.GetApartmentState();
                    dispatchedThreadId = Thread.CurrentThread.ManagedThreadId;
                    receivedSnippet = snippet != null &&
                        (string)snippet.GetType().GetProperty("Name").GetValue(
                            snippet,
                            null) == "Test";
                    actionCompleted.Set();
                };
                actionField.SetValue(addIn, action);
                if ((int)addIn.GetType().GetMethod("GetSnippetCount").Invoke(
                    addIn,
                    new object[] { null }) != 1)
                {
                    return false;
                }
                string selectedItemId =
                    (string)addIn.GetType().GetMethod("GetSnippetId").Invoke(
                        addIn,
                        new object[] { null, 0 });
                MtaInvocation invocation = InvokeOnMta(delegate
                {
                    try
                    {
                        addIn.GetType().GetMethod("InsertGallerySnippet").Invoke(
                            addIn,
                            new object[] { null, selectedItemId, 0 });
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
                    receivedSnippet &&
                    !actionStartedBeforeReturn &&
                    invocation.Duration < TimeSpan.FromSeconds(2) &&
                    apartment == ApartmentState.STA &&
                    dispatchedThreadId != callbackThreadId;
            }
            finally
            {
                if (callbackFinished) actionCompleted.Dispose();
                FieldInfo dispatcherField = addIn.GetType().GetField(
                    "dispatcher",
                    BindingFlags.Instance | BindingFlags.NonPublic);
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
