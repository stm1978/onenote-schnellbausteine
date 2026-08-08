using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Schnellbausteine.AddIn
{
    internal interface IOneNoteWindowActivator
    {
        void Activate(IntPtr windowHandle);
    }

    internal sealed class OneNoteWindowActivationException : Exception
    {
        public OneNoteWindowActivationException(string message)
            : base(message)
        {
        }

        public OneNoteWindowActivationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    internal sealed class OneNoteWindowActivator : IOneNoteWindowActivator
    {
        public void Activate(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero || !NativeMethods.IsWindow(windowHandle))
            {
                throw new OneNoteWindowActivationException(
                    "Das OneNote-Hauptfenster ist nicht mehr verfügbar.");
            }

            IntPtr foregroundWindow = NativeMethods.GetForegroundWindow();
            uint foregroundThread = foregroundWindow == IntPtr.Zero
                ? 0
                : NativeMethods.GetWindowThreadProcessId(foregroundWindow, IntPtr.Zero);
            uint currentThread = NativeMethods.GetCurrentThreadId();
            bool attached = false;

            if (foregroundThread != 0 && foregroundThread != currentThread)
            {
                attached = NativeMethods.AttachThreadInput(
                    foregroundThread,
                    currentThread,
                    true);
                if (!attached)
                {
                    throw new OneNoteWindowActivationException(
                        "Die Eingabewarteschlange konnte nicht mit OneNote verbunden werden.",
                        new Win32Exception(Marshal.GetLastWin32Error()));
                }
            }

            OneNoteWindowActivationException failure = null;
            try
            {
                Log.Trace("onenote.window.SetForegroundWindow", "before");
                if (!NativeMethods.SetForegroundWindow(windowHandle))
                {
                    failure = new OneNoteWindowActivationException(
                        "Windows hat die Aktivierung des OneNote-Hauptfensters abgelehnt.");
                }
                Log.Trace("onenote.window.SetForegroundWindow", "after");
            }
            finally
            {
                if (attached && !NativeMethods.AttachThreadInput(
                    foregroundThread,
                    currentThread,
                    false))
                {
                    var detachFailure = new OneNoteWindowActivationException(
                        "Die gemeinsame Eingabewarteschlange konnte nicht getrennt werden.",
                        new Win32Exception(Marshal.GetLastWin32Error()));
                    if (failure == null)
                    {
                        failure = detachFailure;
                    }
                    else
                    {
                        Log.Error("OneNote-Fokus: Eingabewarteschlange trennen", detachFailure);
                    }
                }
            }

            if (failure != null) throw failure;
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool AttachThreadInput(
                uint idAttach,
                uint idAttachTo,
                bool attach);

            [DllImport("kernel32.dll")]
            internal static extern uint GetCurrentThreadId();

            [DllImport("user32.dll")]
            internal static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            internal static extern uint GetWindowThreadProcessId(
                IntPtr windowHandle,
                IntPtr processId);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsWindow(IntPtr windowHandle);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool SetForegroundWindow(IntPtr windowHandle);
        }
    }
}
