using System;
using System.Reflection;

namespace Schnellbausteine.Tests
{
    internal static class AddInSmoke
    {
        private static int Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.Error.WriteLine("Usage: Schnellbausteine.AddInSmoke.exe <add-in-path>");
                return 1;
            }

            try
            {
                Assembly assembly = Assembly.LoadFrom(args[0]);
                Type type = assembly.GetType("Schnellbausteine.AddIn.AddIn", true);
                object addIn = Activator.CreateInstance(type);
                string ribbonXml = (string)type.GetMethod("GetCustomUI").Invoke(
                    addIn, new object[] { "Microsoft.OneNote" });
                string menuXml = (string)type.GetMethod("GetSnippetMenuContent").Invoke(
                    addIn, new object[] { null });

                if (!ribbonXml.Contains("sqInsertMenu") ||
                    !ribbonXml.Contains("label=\"Einfügen\"") ||
                    !menuXml.Contains("sqEmpty"))
                {
                    Console.Error.WriteLine("Ribbon content is incomplete.");
                    return 1;
                }

                Console.WriteLine("Managed add-in and Ribbon smoke test succeeded.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }
    }
}
