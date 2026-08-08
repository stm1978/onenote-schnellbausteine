using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Schnellbausteine.AddIn
{
    internal interface IOneNoteWindowActivator
    {
        void Activate(IntPtr windowHandle);
    }

    internal interface IOneNoteWindowNative
    {
        bool AttachThreadInput(uint attachThread, uint attachToThread, bool attach);
        bool BringWindowToTop(IntPtr windowHandle);
        uint GetCurrentProcessId();
        uint GetCurrentThreadId();
        IntPtr GetFocus();
        IntPtr GetForegroundWindow();
        uint GetWindowThreadProcessId(IntPtr windowHandle, out uint processId);
        bool IsChild(IntPtr parentWindow, IntPtr windowHandle);
        bool IsIconic(IntPtr windowHandle);
        bool IsWindow(IntPtr windowHandle);
        bool SetForegroundWindow(IntPtr windowHandle);
        bool ShowWindow(IntPtr windowHandle, int command);
    }

    internal sealed class OneNoteWindowActivator : IOneNoteWindowActivator
    {
        private const int RestoreWindow = 9;
        private readonly IOneNoteWindowNative native;

        public OneNoteWindowActivator()
            : this(new Win32OneNoteWindowNative())
        {
        }

        internal OneNoteWindowActivator(IOneNoteWindowNative native)
        {
            if (native == null) throw new ArgumentNullException("native");
            this.native = native;
        }

        public void Activate(IntPtr windowHandle)
        {
            uint currentProcess = native.GetCurrentProcessId();
            uint currentThread = native.GetCurrentThreadId();
            IntPtr foregroundWindow = native.GetForegroundWindow();
            uint foregroundProcess;
            uint foregroundThread = GetWindowThread(foregroundWindow, out foregroundProcess);
            uint targetProcess;
            uint targetThread = GetWindowThread(windowHandle, out targetProcess);

            Log.Diagnostic(
                "onenote.window.identities",
                "currentPid=" + currentProcess.ToString(CultureInfo.InvariantCulture) +
                " currentTid=" + currentThread.ToString(CultureInfo.InvariantCulture) +
                " foregroundHwnd=" + FormatHandle(foregroundWindow) +
                " foregroundPid=" + foregroundProcess.ToString(CultureInfo.InvariantCulture) +
                " foregroundTid=" + foregroundThread.ToString(CultureInfo.InvariantCulture) +
                " targetHwnd=" + FormatHandle(windowHandle) +
                " targetPid=" + targetProcess.ToString(CultureInfo.InvariantCulture) +
                " targetTid=" + targetThread.ToString(CultureInfo.InvariantCulture));

            if (windowHandle == IntPtr.Zero ||
                !native.IsWindow(windowHandle) ||
                targetThread == 0)
            {
                Log.Diagnostic("onenote.window.activation", "validTarget=false");
                return;
            }

            bool foregroundAttached = false;
            bool targetAttached = false;
            try
            {
                if (foregroundThread != 0 && foregroundThread != currentThread)
                {
                    foregroundAttached = native.AttachThreadInput(
                        currentThread,
                        foregroundThread,
                        true);
                    LogBoolean("onenote.window.attachForeground", foregroundAttached);
                }

                if (targetThread != currentThread && targetThread != foregroundThread)
                {
                    targetAttached = native.AttachThreadInput(
                        currentThread,
                        targetThread,
                        true);
                    LogBoolean("onenote.window.attachTarget", targetAttached);
                }

                bool iconic = native.IsIconic(windowHandle);
                LogBoolean("onenote.window.iconic", iconic);
                if (iconic)
                {
                    LogBoolean(
                        "onenote.window.restore",
                        native.ShowWindow(windowHandle, RestoreWindow));
                }

                bool setForeground = native.SetForegroundWindow(windowHandle);
                LogBoolean("onenote.window.setForeground", setForeground);
                LogBoolean(
                    "onenote.window.bringToTop",
                    native.BringWindowToTop(windowHandle));

                IntPtr finalForegroundWindow = native.GetForegroundWindow();
                IntPtr focusedWindow = native.GetFocus();
                bool foregroundMatches = finalForegroundWindow == windowHandle;
                bool editorFocused = focusedWindow != IntPtr.Zero &&
                    focusedWindow != windowHandle &&
                    native.IsChild(windowHandle, focusedWindow);
                Log.Diagnostic(
                    "onenote.window.result",
                    "foregroundHwnd=" + FormatHandle(finalForegroundWindow) +
                    " focusedHwnd=" + FormatHandle(focusedWindow) +
                    " foregroundMatches=" + FormatBoolean(foregroundMatches) +
                    " editorFocused=" + FormatBoolean(editorFocused));
            }
            finally
            {
                if (targetAttached)
                {
                    LogBoolean(
                        "onenote.window.detachTarget",
                        native.AttachThreadInput(currentThread, targetThread, false));
                }
                if (foregroundAttached)
                {
                    LogBoolean(
                        "onenote.window.detachForeground",
                        native.AttachThreadInput(currentThread, foregroundThread, false));
                }
            }
        }

        private uint GetWindowThread(IntPtr windowHandle, out uint processId)
        {
            processId = 0;
            return windowHandle == IntPtr.Zero
                ? 0
                : native.GetWindowThreadProcessId(windowHandle, out processId);
        }

        private static void LogBoolean(string operation, bool value)
        {
            Log.Diagnostic(operation, "result=" + FormatBoolean(value));
        }

        private static string FormatBoolean(bool value)
        {
            return value ? "true" : "false";
        }

        private static string FormatHandle(IntPtr value)
        {
            return "0x" + value.ToInt64().ToString("X", CultureInfo.InvariantCulture);
        }
    }

    internal sealed class Win32OneNoteWindowNative : IOneNoteWindowNative
    {
        public bool AttachThreadInput(uint attachThread, uint attachToThread, bool attach)
        {
            return NativeMethods.AttachThreadInput(attachThread, attachToThread, attach);
        }

        public bool BringWindowToTop(IntPtr windowHandle)
        {
            return NativeMethods.BringWindowToTop(windowHandle);
        }

        public uint GetCurrentProcessId()
        {
            return NativeMethods.GetCurrentProcessId();
        }

        public uint GetCurrentThreadId()
        {
            return NativeMethods.GetCurrentThreadId();
        }

        public IntPtr GetFocus()
        {
            return NativeMethods.GetFocus();
        }

        public IntPtr GetForegroundWindow()
        {
            return NativeMethods.GetForegroundWindow();
        }

        public uint GetWindowThreadProcessId(IntPtr windowHandle, out uint processId)
        {
            return NativeMethods.GetWindowThreadProcessId(windowHandle, out processId);
        }

        public bool IsChild(IntPtr parentWindow, IntPtr windowHandle)
        {
            return NativeMethods.IsChild(parentWindow, windowHandle);
        }

        public bool IsIconic(IntPtr windowHandle)
        {
            return NativeMethods.IsIconic(windowHandle);
        }

        public bool IsWindow(IntPtr windowHandle)
        {
            return NativeMethods.IsWindow(windowHandle);
        }

        public bool SetForegroundWindow(IntPtr windowHandle)
        {
            return NativeMethods.SetForegroundWindow(windowHandle);
        }

        public bool ShowWindow(IntPtr windowHandle, int command)
        {
            return NativeMethods.ShowWindow(windowHandle, command);
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool AttachThreadInput(
                uint idAttach,
                uint idAttachTo,
                bool attach);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool BringWindowToTop(IntPtr windowHandle);

            [DllImport("kernel32.dll")]
            internal static extern uint GetCurrentProcessId();

            [DllImport("kernel32.dll")]
            internal static extern uint GetCurrentThreadId();

            [DllImport("user32.dll")]
            internal static extern IntPtr GetFocus();

            [DllImport("user32.dll")]
            internal static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            internal static extern uint GetWindowThreadProcessId(
                IntPtr windowHandle,
                out uint processId);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsChild(
                IntPtr parentWindow,
                IntPtr windowHandle);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsIconic(IntPtr windowHandle);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsWindow(IntPtr windowHandle);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool SetForegroundWindow(IntPtr windowHandle);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool ShowWindow(IntPtr windowHandle, int command);
        }
    }
}
