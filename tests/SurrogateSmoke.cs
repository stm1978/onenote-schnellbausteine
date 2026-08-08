using System;
using System.Runtime.InteropServices;
using Extensibility;

namespace Schnellbausteine.Tests
{
    internal static class SurrogateSmoke
    {
        private const uint ClsctxLocalServer = 0x4;

        [DllImport("ole32.dll")]
        private static extern int CoCreateInstance(
            ref Guid classId,
            IntPtr outer,
            uint classContext,
            ref Guid interfaceId,
            out IntPtr instance);

        private static int Main()
        {
            Guid classId = new Guid("CCF3B64A-9580-4D44-BDCE-49BFE95AC510");
            Guid extensibilityId = new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744");
            IntPtr pointer;
            int result = CoCreateInstance(
                ref classId,
                IntPtr.Zero,
                ClsctxLocalServer,
                ref extensibilityId,
                out pointer);

            if (result < 0)
            {
                Console.Error.WriteLine("Surrogate activation failed: 0x" + result.ToString("X8"));
                return 1;
            }

            IDTExtensibility2 addIn = null;
            try
            {
                addIn = (IDTExtensibility2)Marshal.GetTypedObjectForIUnknown(
                    pointer, typeof(IDTExtensibility2));
                Array custom = null;
                addIn.OnConnection(
                    null,
                    ext_ConnectMode.ext_cm_Startup,
                    null,
                    ref custom);
                addIn.OnStartupComplete(ref custom);
                addIn.OnBeginShutdown(ref custom);
                addIn.OnDisconnection(
                    ext_DisconnectMode.ext_dm_HostShutdown,
                    ref custom);
            }
            finally
            {
                if (addIn != null && Marshal.IsComObject(addIn))
                {
                    Marshal.FinalReleaseComObject(addIn);
                }
                Marshal.Release(pointer);
            }

            Console.WriteLine("COM surrogate lifecycle succeeded.");
            return 0;
        }
    }
}
