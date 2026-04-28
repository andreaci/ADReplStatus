using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADReplStatus
{
    internal class SettingsManager
    {
        public static ADHelper ADHelperObject { get; set; } = new ADHelper();

        private static bool __darkMode = false;
        public static bool gDarkMode { get { return __darkMode; } set { WriteRegKey("DarkMode", value); __darkMode = value; } }

        private static void WriteRegKey(string keyName, bool value)
        {
            var key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\ADREPLSTATUS", true);

            if (key != null)
            {
                key.SetValue(keyName, value ? 1 : 0);
                key.Dispose();
            }
        }

        public static bool gErrorsOnly { get; set; } = false;

        internal static void LoadSettings()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\ADREPLSTATUS", false))
                {
                    if (key != null)
                    {
                        ADHelperObject.gForestName = key.GetValue("ForestName", string.Empty).ToString();
                        gDarkMode = Convert.ToBoolean(key.GetValue("DarkMode", false));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occured while trying to read app settings from the registry!\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ADHelperObject.Init();

        
        }
    }
}
