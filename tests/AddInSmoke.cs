using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
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

        private sealed class RibbonControlProxy : RealProxy
        {
            private readonly string id;

            public RibbonControlProxy(Type interfaceType, string id)
                : base(interfaceType)
            {
                this.id = id;
            }

            public override IMessage Invoke(IMessage message)
            {
                var call = (IMethodCallMessage)message;
                object value = call.MethodName == "get_Id" ? id : null;
                return new ReturnMessage(value, null, 0, call.LogicalCallContext, call);
            }
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
            string menuRepositoryRoot = Path.Combine(
                Path.GetTempPath(),
                "Schnellbausteine.AddInSmoke",
                Guid.NewGuid().ToString("N"));
            string manageRepositoryRoot = Path.Combine(
                Path.GetTempPath(),
                "Schnellbausteine.ManageDialogSmoke",
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
                MethodInfo insertSnippetMethod = type.GetMethod("InsertSnippet");
                MethodInfo manageSnippetsMethod = type.GetMethod("ManageSnippets");
                MethodInfo insertSnippetCoreMethod = type.GetMethod(
                    "InsertSnippetCore",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo manageSnippetsCoreMethod = type.GetMethod(
                    "ManageSnippetsCore",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo invalidateMenuMethod = type.GetMethod(
                    "InvalidateSnippetMenu",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                MethodInfo readCurrentPageMethod = clientType.GetMethod("ReadCurrentPage");
                MethodInfo menuContentMethod = type.GetMethod("GetSnippetMenuContent");
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
                Directory.CreateDirectory(menuRepositoryRoot);
                object menuRepository = Activator.CreateInstance(
                    repositoryType,
                    new object[] { menuRepositoryRoot });
                XDocument menuFragment = XDocument.Parse(
                    "<sq:Snippet xmlns:sq=\"urn:schnellbausteine:v1\" " +
                    "xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" " +
                    "version=\"1\"><sq:TagDefs/><sq:QuickStyleDefs/><sq:Content>" +
                    "<one:OE><one:T><![CDATA[Test]]></one:T></one:OE>" +
                    "</sq:Content></sq:Snippet>");
                repositoryType.GetMethod("Save").Invoke(
                    menuRepository,
                    new object[] { "Test & <XML>", menuFragment, false });
                ConstructorInfo repositoryConstructor = type.GetConstructor(
                    BindingFlags.Instance | BindingFlags.NonPublic,
                    null,
                    new[] { repositoryType },
                    null);
                object menuAddIn = repositoryConstructor.Invoke(
                    new[] { menuRepository });
                Type ribbonControlType = menuContentMethod.GetParameters()[0].ParameterType;
                object mainMenuControl = new RibbonControlProxy(
                    ribbonControlType,
                    "sqMain").GetTransparentProxy();
                string menuContentXml = (string)menuContentMethod.Invoke(
                    menuAddIn, new[] { mainMenuControl });
                bool itemCallbackHarnessWorks = VerifyDeferredMenuCallback(
                    menuAddIn,
                    insertSnippetActionField,
                    XDocument.Parse(menuContentXml));
                bool manageDialogRefreshWorks = VerifyManageDialogRefresh(
                    manageDialogType,
                    repositoryType,
                    manageRepositoryRoot);
                string ribbonXml = (string)type.GetMethod("GetCustomUI").Invoke(
                    addIn, new object[] { "Microsoft.OneNote" });
                XNamespace ribbonNamespace =
                    "http://schemas.microsoft.com/office/2009/07/customui";
                XDocument ribbon = XDocument.Parse(ribbonXml);
                XElement ribbonGroup = ribbon.Descendants(ribbonNamespace + "group")
                    .SingleOrDefault(element => (string)element.Attribute("id") == "sqGroup");
                XElement insertTab = ribbon.Descendants(ribbonNamespace + "tab")
                    .SingleOrDefault(element => (string)element.Attribute("idMso") == "TabInsert");
                XElement mainMenu = ribbonGroup == null
                    ? null
                    : ribbonGroup.Elements(ribbonNamespace + "dynamicMenu").SingleOrDefault();
                XDocument menuContent = XDocument.Parse(menuContentXml);
                XElement[] menuItems = menuContent.Root.Elements().ToArray();
                XElement snippetButton = menuItems.FirstOrDefault();
                XElement separator = menuItems.Skip(1).FirstOrDefault();
                XElement manageButton = menuItems.Skip(2).FirstOrDefault();
                XElement contextMenus = ribbon.Root.Element(ribbonNamespace + "contextMenus");
                string[] contextMenuIds =
                {
                    "ContextMenuText",
                    "ContextMenuOutlineElementHandle",
                    "ContextMenuInkDrawing",
                    "ContextMenuLassoSelection"
                };
                string[] saveButtonIds =
                {
                    "sqSaveSelectionText",
                    "sqSaveSelectionOutline",
                    "sqSaveSelectionInk",
                    "sqSaveSelectionLasso"
                };
                string[] insertMenuIds =
                {
                    "sqInsertText",
                    "sqInsertOutline",
                    "sqInsertInk",
                    "sqInsertLasso"
                };
                XElement[] contextMenuElements = contextMenus == null
                    ? new XElement[0]
                    : contextMenuIds.Select(id => contextMenus
                        .Elements(ribbonNamespace + "contextMenu")
                        .SingleOrDefault(element => (string)element.Attribute("idMso") == id))
                        .ToArray();
                XElement[] saveButtons = contextMenuElements
                    .Select((menuElement, index) => menuElement == null
                        ? null
                        : menuElement.Elements(ribbonNamespace + "button").SingleOrDefault(
                            element => (string)element.Attribute("id") == saveButtonIds[index]))
                    .ToArray();
                XElement[] insertMenus = contextMenuElements
                    .Select((menuElement, index) => menuElement == null
                        ? null
                        : menuElement.Elements(ribbonNamespace + "dynamicMenu")
                            .SingleOrDefault(element =>
                                (string)element.Attribute("id") == insertMenuIds[index]))
                    .ToArray();
                XElement[] contextSeparators = contextMenuElements
                    .Select(menuElement => menuElement == null
                        ? null
                        : menuElement.Elements(ribbonNamespace + "menuSeparator")
                            .SingleOrDefault())
                    .ToArray();

                if (insertTab == null ||
                    ribbon.Descendants(ribbonNamespace + "tab")
                        .Any(element => (string)element.Attribute("idMso") == "TabHome") ||
                    mainMenu == null ||
                    ribbonGroup.Elements().Count() != 1 ||
                    (string)mainMenu.Attribute("id") != "sqMain" ||
                    (string)mainMenu.Attribute("label") != "Schnellbausteine" ||
                    (string)mainMenu.Attribute("size") != "large" ||
                    (string)mainMenu.Attribute("getContent") != "GetSnippetMenuContent" ||
                    mainMenu.Elements().Any() ||
                    menuContent.Root.Name != ribbonNamespace + "menu" ||
                    menuItems.Length != 3 ||
                    snippetButton.Name != ribbonNamespace + "button" ||
                    (string)snippetButton.Attribute("label") != "Test & <XML>" ||
                    !((string)snippetButton.Attribute("id")).StartsWith("sqSnippet_") ||
                    ((string)snippetButton.Attribute("id")).Contains("Test") ||
                    (string)snippetButton.Attribute("onAction") != "InsertSnippet" ||
                    separator == null ||
                    separator.Name != ribbonNamespace + "menuSeparator" ||
                    !((string)separator.Attribute("id"))
                        .StartsWith("sqManageSeparator_") ||
                    manageButton == null ||
                    manageButton.Name != ribbonNamespace + "button" ||
                    !((string)manageButton.Attribute("id")).StartsWith("sqManage_") ||
                    (string)manageButton.Attribute("label") != "Schnellbausteine verwalten" ||
                    (string)manageButton.Attribute("onAction") != "ManageSnippets" ||
                    ribbon.Descendants(ribbonNamespace + "gallery").Any() ||
                    ribbon.Descendants(ribbonNamespace + "menu").Any() ||
                    ribbon.Descendants(ribbonNamespace + "form").Any() ||
                    contextMenuElements.Any(element => element == null) ||
                    contextMenuElements.Any(element => element.Elements().Count() != 3) ||
                    insertMenus.Any(menuElement => menuElement == null) ||
                    insertMenus.Any(menuElement =>
                        (string)menuElement.Attribute("label") != "Schnellbausteine" ||
                        (string)menuElement.Attribute("imageMso") !=
                            "QuickPartsInsertGallery" ||
                        (string)menuElement.Attribute("getContent") !=
                            "GetSnippetMenuContent" ||
                        menuElement.Elements().Any()) ||
                    contextSeparators.Any(separatorElement => separatorElement == null) ||
                    saveButtons.Any(button => button == null) ||
                    saveButtons.Any(button =>
                        (string)button.Attribute("label") != "Als Schnellbaustein speichern..." ||
                        (string)button.Attribute("imageMso") !=
                            "SaveSelectionToQuickPartGallery" ||
                        (string)button.Attribute("onAction") != "SaveSelection" ||
                        button.Attribute("getVisible") != null ||
                        button.Attribute("getEnabled") != null) ||
                    (string)saveButtons[0].Attribute("insertBeforeMso") != "Cut" ||
                    (string)insertMenus[0].Attribute("insertBeforeMso") != "Cut" ||
                    (string)contextSeparators[0].Attribute("insertBeforeMso") != "Cut" ||
                    ribbon.Descendants()
                        .Where(element => element.Attribute("id") != null)
                        .GroupBy(element => (string)element.Attribute("id"))
                        .Any(group => group.Count() != 1) ||
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
                    menuContentMethod == null ||
                    menuContentMethod.ReturnType != typeof(string) ||
                    menuContentMethod.GetParameters().Length != 1 ||
                    menuContentMethod.GetParameters()[0].ParameterType.FullName !=
                        "Microsoft.Office.Core.IRibbonControl" ||
                    dispatcherType.GetMethod("Post") == null ||
                    dispatcherType.GetMethod("PostAfterCallback") == null ||
                    dispatcherType.GetMethod("Shutdown") == null ||
                    !dispatcherWorks ||
                    !timeoutGuardWorks ||
                    !saveCallbackHarnessWorks ||
                    !manageCallbackHarnessWorks ||
                    !itemCallbackHarnessWorks ||
                    !manageDialogRefreshWorks ||
                    dispatchMethod == null ||
                    !CallsMember(saveSelectionMethod, dispatchMethod) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => CallsMember(saveSelectionMethod, constructor)) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => CallsMember(menuContentMethod, constructor)) ||
                    insertSnippetMethod == null ||
                    insertSnippetMethod.ReturnType != typeof(void) ||
                    insertSnippetMethod.GetParameters().Length != 1 ||
                    insertSnippetMethod.GetParameters()[0].ParameterType.FullName !=
                        "Microsoft.Office.Core.IRibbonControl" ||
                    !CallsMethod(insertSnippetMethod, dispatchMethod) ||
                    forbiddenCallbackConstructors.Any(
                        constructor => CallsMember(insertSnippetMethod, constructor)) ||
                    insertSnippetCoreMethod == null ||
                    !CallsMethod(insertSnippetCoreMethod, readCurrentPageMethod) ||
                    manageSnippetsCoreMethod == null ||
                    invalidateMenuMethod == null ||
                    !CallsMember(saveSelectionMethod.DeclaringType.GetMethod(
                        "SaveSelectionCore",
                        BindingFlags.Instance | BindingFlags.NonPublic),
                        invalidateMenuMethod) ||
                    !CallsMember(manageSnippetsCoreMethod, invalidateMenuMethod) ||
                    !CallsMember(
                        invalidateMenuMethod,
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
                    type.GetMethod("CanSaveSelection") != null ||
                    selectionMethod == null ||
                    type.GetMethod("SaveSelection") == null ||
                    type.GetMethod("GetSnippetCount") != null ||
                    type.GetMethod("GetSnippetLabel") != null ||
                    type.GetMethod("GetSnippetId") != null ||
                    type.GetMethod("GetSnippetGalleryEnabled") != null ||
                    type.GetMethod("InsertGallerySnippet") != null)
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
                if (Directory.Exists(menuRepositoryRoot))
                {
                    Directory.Delete(menuRepositoryRoot, true);
                }
                if (Directory.Exists(manageRepositoryRoot))
                {
                    Directory.Delete(manageRepositoryRoot, true);
                }
            }
        }

        private static bool VerifyManageDialogRefresh(
            Type manageDialogType,
            Type repositoryType,
            string repositoryRoot)
        {
            Directory.CreateDirectory(repositoryRoot);
            object repository = Activator.CreateInstance(
                repositoryType,
                new object[] { repositoryRoot });
            MethodInfo save = repositoryType.GetMethod("Save");
            MethodInfo rename = repositoryType.GetMethod("Rename");
            MethodInfo delete = repositoryType.GetMethod("Delete");
            XDocument fragment = XDocument.Parse(
                "<sq:Snippet xmlns:sq=\"urn:schnellbausteine:v1\" " +
                "xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" " +
                "version=\"1\"><sq:TagDefs/><sq:QuickStyleDefs/><sq:Content>" +
                "<one:OE><one:T><![CDATA[Test]]></one:T></one:OE>" +
                "</sq:Content></sq:Snippet>");
            object alpha = save.Invoke(
                repository,
                new object[] { "Alpha", fragment, false });
            object beta = save.Invoke(
                repository,
                new object[] { "Beta", fragment, false });
            PropertyInfo id = alpha.GetType().GetProperty("Id");
            string alphaId = (string)id.GetValue(alpha, null);
            string betaId = (string)id.GetValue(beta, null);
            Form dialog = null;
            try
            {
                dialog = (Form)Activator.CreateInstance(
                    manageDialogType,
                    new object[] { repository, new Action(delegate { }) });
                ListBox list = (ListBox)manageDialogType.GetField(
                    "list",
                    BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dialog);
                MethodInfo loadItems = manageDialogType.GetMethod(
                    "LoadItems",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (ListText(list) != "Alpha,Beta") return false;

                list.SelectedItem = list.Items.Cast<object>().Single(item =>
                    (string)id.GetValue(item, null) == alphaId);
                delete.Invoke(repository, new object[] { betaId });
                loadItems.Invoke(dialog, null);
                if (ListText(list) != "Alpha" ||
                    (string)id.GetValue(list.SelectedItem, null) != alphaId)
                {
                    return false;
                }

                rename.Invoke(repository, new object[] { alphaId, "Aaron" });
                loadItems.Invoke(dialog, null);
                if (ListText(list) != "Aaron" ||
                    (string)id.GetValue(list.SelectedItem, null) != alphaId)
                {
                    return false;
                }

                object zulu = save.Invoke(
                    repository,
                    new object[] { "Zulu", fragment, false });
                string zuluId = (string)id.GetValue(zulu, null);
                loadItems.Invoke(dialog, null);
                if (ListText(list) != "Aaron,Zulu" ||
                    (string)id.GetValue(list.SelectedItem, null) != alphaId)
                {
                    return false;
                }

                delete.Invoke(repository, new object[] { alphaId });
                loadItems.Invoke(dialog, null);
                if (ListText(list) != "Zulu" ||
                    list.SelectedItem == null ||
                    (string)id.GetValue(list.SelectedItem, null) != zuluId)
                {
                    return false;
                }

                delete.Invoke(repository, new object[] { zuluId });
                loadItems.Invoke(dialog, null);
                if (list.Items.Count != 0 || list.SelectedItem != null) return false;

                save.Invoke(repository, new object[] { "Gamma", fragment, false });
                loadItems.Invoke(dialog, null);
                return ListText(list) == "Gamma";
            }
            finally
            {
                if (dialog != null) dialog.Dispose();
            }
        }

        private static string ListText(ListBox list)
        {
            return string.Join(
                ",",
                list.Items.Cast<object>().Select(item => list.GetItemText(item)));
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

        private static bool VerifyDeferredMenuCallback(
            object addIn,
            FieldInfo actionField,
            XDocument menuContent)
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
                            null) == "Test & <XML>";
                    actionCompleted.Set();
                };
                actionField.SetValue(addIn, action);
                string selectedItemId = (string)menuContent.Root.Elements().First()
                    .Attribute("id");
                Type ribbonControlType = addIn.GetType().GetMethod("InsertSnippet")
                    .GetParameters()[0].ParameterType;
                object control = new RibbonControlProxy(
                    ribbonControlType,
                    selectedItemId).GetTransparentProxy();
                MtaInvocation invocation = InvokeOnMta(delegate
                {
                    try
                    {
                        addIn.GetType().GetMethod("InsertSnippet").Invoke(
                            addIn,
                            new[] { control });
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
