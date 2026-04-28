using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADReplStatus
{
    internal class LoggingManager
    {   
        public static string gLogfileName { get; set; } = string.Empty;

        private static bool __enabled = false;
        public static bool gLoggingEnabled
        {
            get
            {
                return __enabled;
            }
            set
            {
                __enabled = value;
                if (__enabled)
                {
                    gLogfileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"ADReplStatus_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                    AppendText("Logging started.");
                }
                else
                    AppendText("Logging stopped.");
                
            }
        }

        public static void AppendText(string text) {

            if (gLoggingEnabled)
            {
                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] {text}\n");
            }
        }

        internal static void Error(string errorMessage, bool showMsgBox = false)
        {
            AppendText("ERROR: " + errorMessage);
            new Thread(() => MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error));
        }
    }
}
