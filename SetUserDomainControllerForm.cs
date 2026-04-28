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
    public partial class SetUserDomainControllerForm : Form
    {
        public SetUserDomainControllerForm()
        {
            InitializeComponent();
        }

        private void SetForestNameButton_Click(object sender, EventArgs e)
        {

            if (ADReplStatusForm.gUseUserDomainController)
            {
                //The user cleared out the input box and clicked set
                if(SetUserDomainControllerTextBox.Text.Length < 1)
                {
                    LoggingManager.AppendText("Clearing user specified domain controller and disabling global.");
                    ADReplStatusForm.gUseUserDomainController = false;
                }
                else
                {
                    LoggingManager.AppendText($"Changing user specified domain controller to {SetUserDomainControllerTextBox.Text}");
                    ADReplStatusForm.gUserDomainController = SetUserDomainControllerTextBox.Text;
                }

                this.Dispose();
                return;
            }

            LoggingManager.AppendText($"Setting user specified domain controller to {SetUserDomainControllerTextBox.Text} and enabling global.");
            
            ADReplStatusForm.gUseUserDomainController = true;
            ADReplStatusForm.gUserDomainController = SetUserDomainControllerTextBox.Text;

            this.Dispose();
            return;
        }

        private void SetUserDomainControllerForm_Load(object sender, EventArgs e)
        {
            SetUserDomainControllerTextBox.Text = string.Empty;

            if (ADReplStatusForm.gUseUserDomainController)
                SetUserDomainControllerTextBox.Text = ADReplStatusForm.gUserDomainController;
        }

        private void SetUserDomainControllerTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SetForestNameButton_Click(this, new EventArgs());
            }
        }
    }
}
