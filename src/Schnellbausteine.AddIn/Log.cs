using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Security;
using System.Text;
using System.Threading;

namespace Schnellbausteine.AddIn
{
    internal static class Log
    {
        private static readonly ConcurrentQueue<LogEntry> Entries =
            new ConcurrentQueue<LogEntry>();
        private static readonly AutoResetEvent Pending = new AutoResetEvent(false);
        private static readonly string PathName = GetPathName();
        private static readonly int ProcessId = Process.GetCurrentProcess().Id;
        private static int writerStarted;

        private sealed class LogEntry
        {
            public string Text;
            public ManualResetEventSlim Completion;
        }

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

        public static bool Flush(TimeSpan timeout)
        {
            EnsureWriter();
            var completion = new ManualResetEventSlim(false);
            Entries.Enqueue(new LogEntry { Completion = completion });
            Pending.Set();
            bool completed = completion.Wait(timeout);
            if (completed) completion.Dispose();
            return completed;
        }

        private static void Write(string operation, string phase, string detail)
        {
            var line = new StringBuilder()
                .Append(DateTime.Now.ToString("o"))
                .Append(" pid=").Append(ProcessId)
                .Append(" thread=").Append(Thread.CurrentThread.ManagedThreadId)
                .Append(" apartment=").Append(Thread.CurrentThread.GetApartmentState())
                .Append(" operation=").Append(operation)
                .Append(" phase=").Append(phase)
                .AppendLine();
            if (!string.IsNullOrEmpty(detail)) line.AppendLine(detail);

            EnsureWriter();
            Entries.Enqueue(new LogEntry { Text = line.ToString() });
            Pending.Set();
        }

        private static void EnsureWriter()
        {
            if (Interlocked.CompareExchange(ref writerStarted, 1, 0) != 0) return;
            var writer = new Thread(WriterLoop);
            writer.IsBackground = true;
            writer.Name = "Schnellbausteine Log";
            writer.SetApartmentState(ApartmentState.MTA);
            writer.Start();
        }

        private static void WriterLoop()
        {
            while (true)
            {
                Pending.WaitOne();
                LogEntry entry;
                while (Entries.TryDequeue(out entry))
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(entry.Text))
                        {
                            string directory = Path.GetDirectoryName(PathName);
                            if (!string.IsNullOrEmpty(directory))
                            {
                                Directory.CreateDirectory(directory);
                            }
                            File.AppendAllText(
                                PathName, entry.Text, new UTF8Encoding(false));
                        }
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
                    finally
                    {
                        if (entry.Completion != null) entry.Completion.Set();
                    }
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
