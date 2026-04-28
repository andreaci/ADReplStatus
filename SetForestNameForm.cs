using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADReplStatus
{
    public partial class SetForestNameForm : Form
    {
        public SetForestNameForm()
        {
            InitializeComponent();
        }

        private void SetForestNameForm_Load(object sender, EventArgs e)
        {
            if (SettingsManager.gDarkMode == true)
            {
                this.BackColor = Color.FromArgb(32, 32, 32);
                EnterForestNameLabel.BackColor = Color.FromArgb(32, 32, 32);
                EnterForestNameLabel.ForeColor = Color.White;
                SetForestNameTextBox.BackColor = Color.FromArgb(32, 32, 32);
                SetForestNameTextBox.ForeColor = Color.White;
                SetForestNameButton.BackColor = Color.FromArgb(32, 32, 32);
                SetForestNameButton.ForeColor = Color.White;
                SaveForestCheckBox.ForeColor = Color.White;
            }
        }

        private void SetForestNameButton_Click(object sender, EventArgs e)
        {
            if (SetForestNameTextBox.Text.Length > 0)
            {
                SettingsManager.ADHelperObject.gForestName = SetForestNameTextBox.Text;
                if (SaveForestCheckBox.Checked)
                    SettingsManager.WriteRegKey("ForestName", SetForestNameTextBox.Text);

                LoggingManager.AppendText($"Forest name set to: {SettingsManager.ADHelperObject.gForestName}\n");
                this.Dispose();
            }
        }
    }
}
