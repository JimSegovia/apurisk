using System;
using System.IO;

namespace Apurisk.ExcelAddIn.Diagnostics
{
    internal static class AddInLog
    {
        private static readonly string LogPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Apurisk.AddIn.log");

        public static void Write(string message)
        {
            try
            {
                File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        public static void WriteException(string context, Exception exception)
        {
            Write(context + " | " + exception.GetType().FullName + " | " + exception.Message);
        }
    }
}
