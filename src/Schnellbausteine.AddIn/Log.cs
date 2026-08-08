using System;
using System.Diagnostics;
using System.IO;
using System.Security;
using System.Text;
using System.Threading;

namespace Schnellbausteine.AddIn
{
    internal static class Log
    {
        private static readonly object Gate = new object();
        private static readonly string PathName = GetPathName();

        public static void Error(string operation, Exception exception)
        {
            Write(operation, "error", exception.ToString());
        }

        public static void Info(string message)
        {
            Write("info", message, null);
        }

        public static void Trace(string operation, string phase)
        {
            Write(operation, phase, null);
        }

        private static void Write(string operation, string phase, string detail)
        {
            lock (Gate)
            {
                try
                {
                    string directory = Path.GetDirectoryName(PathName);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                    var line = new StringBuilder()
                        .Append(DateTime.Now.ToString("o"))
                        .Append(" pid=").Append(Process.GetCurrentProcess().Id)
                        .Append(" thread=").Append(Thread.CurrentThread.ManagedThreadId)
                        .Append(" apartment=").Append(Thread.CurrentThread.GetApartmentState())
                        .Append(" operation=").Append(operation)
                        .Append(" phase=").Append(phase)
                        .AppendLine();
                    if (!string.IsNullOrEmpty(detail)) line.AppendLine(detail);
                    File.AppendAllText(PathName, line.ToString(), new UTF8Encoding(false));
                }
                catch (IOException exception)
                {
                    Debug.WriteLine(exception);
                }
                catch (UnauthorizedAccessException exception)
                {
                    Debug.WriteLine(exception);
                }
                catch (SecurityException exception)
                {
                    Debug.WriteLine(exception);
                }
                catch (ArgumentException exception)
                {
                    Debug.WriteLine(exception);
                }
                catch (NotSupportedException exception)
                {
                    Debug.WriteLine(exception);
                }
            }
        }

        private static string GetPathName()
        {
            string overridePath = Environment.GetEnvironmentVariable(
                "SCHNELLBAUSTEINE_LOG_PATH");
            if (!string.IsNullOrWhiteSpace(overridePath)) return overridePath;
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Schnellbausteine", "addin.log");
        }
    }
}
