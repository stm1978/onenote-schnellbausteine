using System;
using System.IO;
using System.Text;

namespace Schnellbausteine.AddIn
{
    internal static class Log
    {
        private static readonly object Gate = new object();
        private static readonly string PathName = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Schnellbausteine", "addin.log");

        public static void Error(string operation, Exception exception)
        {
            lock (Gate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(PathName));
                File.AppendAllText(PathName,
                    DateTime.Now.ToString("s") + " " + operation + Environment.NewLine + exception + Environment.NewLine,
                    new UTF8Encoding(false));
            }
        }
    }
}
