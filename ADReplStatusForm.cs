using BrightIdeasSoftware;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADReplStatus
{
    public partial class ADReplStatusForm : Form
    {

        public static string gPassword = string.Empty;

        public static string gTarget = string.Empty;

        public static string gUserDomainController = string.Empty;

        public static string gUsername = string.Empty;

        public static bool gUseUserDomainController;


        public ADReplStatusForm()
        {
            InitializeComponent();

            DiscoveredDCsUpdated += SyncDiscoveredDCs;
        }

        public event Action DiscoveredDCsUpdated;

        private void ADReplStatusForm_Load(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(RefreshButton, "Refresh Replication Status");
            toolTip1.SetToolTip(EnableLoggingButton, "Enable Logging");
            toolTip1.SetToolTip(SetForestButton, "Manually Set Forest");
            toolTip1.SetToolTip(AlternateCredsButton, "Provide Alternate Credentials");
            toolTip1.SetToolTip(ErrorsOnlyButton, "Show Errors Only");

            SettingsManager.LoadSettings();
            UpdateForm();
        }

        private void ADReplStatusForm_Resize(object sender, EventArgs e)
        {
            treeListView1.Top = 68;
            treeListView1.Left = 12;

            if (ActiveForm != null)
            {
                treeListView1.Width = ActiveForm.Width - 40;
                treeListView1.Height = ActiveForm.Height - 119;
            }
        }

        private void AlternateCredsButton_Click(object sender, EventArgs e)
        {
            AlternateCredsForm alternateCredsForm = new AlternateCredsForm();

            LoggingManager.AppendText($"AlternateCreds button was clicked.");

            alternateCredsForm.ShowDialog();
            alternateCredsForm.Dispose();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            SettingsManager.ADHelperObject.DoWork(backgroundWorker1);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            LoggingManager.AppendText(e.UserState.ToString());

            if (e.UserState.ToString().StartsWith("ERROR:"))
            {
                MessageBox.Show(e.UserState.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.UserState.ToString().Equals("UPDATEPERCENT"))
            {
                ProgressPercentLabel.Text = $"{e.ProgressPercentage}%";
            }
            else if (e.UserState.ToString().Equals("OnDiscoveredDCsUpdated"))
            {
                OnDiscoveredDCsUpdated();
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressPercentLabel.Visible = false;

            foreach (var control in Controls)
            {
                if (control is Button button)
                {
                    button.Enabled = true;
                }
            }
        }

        private void DarkModeButton_Click(object sender, EventArgs e)
        {
            SettingsManager.gDarkMode = !SettingsManager.gDarkMode;
            UpdateForm();
        }

        private void UpdateForm()
        {
            if (SettingsManager.gDarkMode)
                SetDarkMode();
            else
                SetLightMode();
        }

        private void DCNameColumn_RightClick(object sender, CellRightClickEventArgs e)
        {
            try
            {
                //Only display the menu in the context of the "DC Name" column
                if (e.Column.Text == "DC Name")
                {
                    //Only display the menu if the cell is populated
                    if (treeListView1.SelectedItem.Text != "")
                    {
                        ContextMenuStrip diagnosticMenu = new ContextMenuStrip();
                        diagnosticMenu.ItemClicked += DiagnosticMenuSelector;

                        ObjectListView olv = e.ListView;
                        ToolStripMenuItem pingMenuItem = new ToolStripMenuItem(String.Format("Ping"));
                        diagnosticMenu.Items.Add(pingMenuItem);
                        
                        ToolStripMenuItem rdpMenuItem = new ToolStripMenuItem(String.Format("Initiate RDP connection"));
                        diagnosticMenu.Items.Add(rdpMenuItem);

                        ToolStripMenuItem enterPSSessionMenuItem = new ToolStripMenuItem(String.Format("Enter PowerShell session"));
                        diagnosticMenu.Items.Add(enterPSSessionMenuItem);

                        ToolStripMenuItem portTesterMenuItem = new ToolStripMenuItem(String.Format("Port Tester"));
                        diagnosticMenu.Items.Add(portTesterMenuItem);

                        e.MenuStrip = diagnosticMenu;
                    }
                }
            }
            catch
            {
                //Do nothing, the user simply right-clicked somewhere else, this is the handler ONLY when the selected column is "DC name"
            }
        }

        private void DiagnosticMenuSelector(object sender, ToolStripItemClickedEventArgs e)
        {
            switch (e.ClickedItem.ToString())
            {
                case "Ping":
                    LoggingManager.AppendText("Diagnostic ping menu opened.");
                    DiagnosticPing(sender, e);
                    break;

                case "Initiate RDP connection":
                    DiagnosticRdp(sender, e);
                    break;

                case "Enter PowerShell session":
                    DiagnosticPSSession(sender, e);
                    break;

                case "Port Tester":
                    DiagnosticNetworkTester(sender, e);
                    break;
            }
        }

        private void DiagnosticNetworkTester(object sender, ToolStripItemClickedEventArgs e)
        {
            gTarget = treeListView1.SelectedItem.Text;

            PortTester protocolTesterForm = new PortTester();
            LoggingManager.AppendText("Port Tester button was clicked.");

            protocolTesterForm.ShowDialog();

            protocolTesterForm.Dispose();
        }

        public void DiagnosticPing(object sender, ToolStripItemClickedEventArgs e)
        {
            string destination = treeListView1.SelectedItem.Text;

            if (destination != "")
            {
                using (var dialog = new Form())
                {
                    //Set up the ping test window
                    dialog.Text = "Ping Test";
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.MaximizeBox = false;
                    dialog.MinimizeBox = false;
                    dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                    dialog.ShowInTaskbar = false;
                    dialog.Width = 290;
                    dialog.Height = 150;

                    var ipv4Button = new Button
                    {
                        Text = "IPv4",
                        Location = new Point(10, 20)
                    };
                    ipv4Button.Click += async (s, ev) => await RunPing(destination, AddressFamily.InterNetwork, dialog);

                    var ipv6Button = new Button
                    {
                        Text = "IPv6",
                        Location = new Point(180, 20)
                    };
                    ipv6Button.Click += async (s, ev) => await RunPing(destination, AddressFamily.InterNetworkV6, dialog);

                    var statusTextBox = new TextBox
                    {
                        Multiline = true,
                        ReadOnly = true,
                        Location = new Point(10, 60),
                        Width = dialog.Width - 45,
                        Height = dialog.Height - 110,
                        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right
                    };

                    dialog.Controls.Add(ipv4Button);
                    dialog.Controls.Add(ipv6Button);
                    dialog.Controls.Add(statusTextBox);

                    //Add support for dark mode
                    if (SettingsManager.gDarkMode)
                    {
                        dialog.BackColor = Color.FromArgb(32, 32, 32);
                        foreach (var control in dialog.Controls)
                        {
                            if (control is Label label)
                            {
                                label.BackColor = Color.FromArgb(32, 32, 32);
                                label.ForeColor = Color.White;
                            }
                            else if (control is TextBox textBox)
                            {
                                textBox.BackColor = Color.FromArgb(32, 32, 32);
                                textBox.ForeColor = Color.White;
                            }
                            else if (control is Button button)
                            {
                                button.BackColor = Color.FromArgb(32, 32, 32);
                                button.ForeColor = Color.White;
                            }
                            else if (control is CheckBox checkBox)
                            {
                                checkBox.BackColor = Color.FromArgb(32, 32, 32);
                                checkBox.ForeColor = Color.White;
                            }
                            else if (control is RadioButton radioButton)
                            {
                                radioButton.BackColor = Color.FromArgb(32, 32, 32);
                                radioButton.ForeColor = Color.White;
                            }
                            else if (control is ListBox listBox)
                            {
                                listBox.BackColor = Color.FromArgb(32, 32, 32);
                                listBox.ForeColor = Color.White;
                            }
                        }
                    }

                    dialog.ShowDialog(this);
                }
            }
        }

        private void DiagnosticPSSession(object sender, ToolStripItemClickedEventArgs e)
        {
            try
            {
                LoggingManager.AppendText($"Initiating remote powershell session to {treeListView1.SelectedItem.Text}");
                string powershellArgs = $"-NoExit $Cred = Get-Credential;Enter-PSSession -ComputerName {treeListView1.SelectedItem.Text} -Credential $Cred";
                Process.Start("powershell.exe", powershellArgs);
            }
            catch (Exception ex)
            {
                string errorMessage = $"ERROR: Enter-PsSession -ComputerName {treeListView1.SelectedItem.Text} failed!\n{ex.Message}\n";
                LoggingManager.Error(errorMessage, true);
                
            }
        }

        private void DiagnosticRdp(object sender, ToolStripItemClickedEventArgs e)
        {
            try
            {
                LoggingManager.AppendText($"Initiating RDP connection to {treeListView1.SelectedItem.Text}");
                string args = $"/v {treeListView1.SelectedItem.Text}";
                Process.Start("mstsc.exe", args);
            }
            catch (Exception ex)
            {
                string errorMessage = $"ERROR: RDP to {treeListView1.SelectedItem.Text} failed!\n{ex.Message}\n";
                LoggingManager.Error(errorMessage, true);
            }
        }

        private void EnableLoggingButton_Click(object sender, EventArgs e)
        {
            LoggingManager.gLoggingEnabled = !LoggingManager.gLoggingEnabled;

            if (LoggingManager.gLoggingEnabled)
            {
                toolTip1.SetToolTip(EnableLoggingButton, "Disable Logging");
                EnableLoggingButton.BackColor = SystemColors.ControlDark;
            }
            else
            {
                toolTip1.SetToolTip(EnableLoggingButton, "Enable Logging");

                if (SettingsManager.gDarkMode)
                    EnableLoggingButton.BackColor = Color.FromArgb(32, 32, 32);
                else
                    EnableLoggingButton.BackColor = SystemColors.Control;

            }
        }

        private void ErrorsOnlyButton_Click(object sender, EventArgs e)
        {
            SettingsManager.gErrorsOnly = !SettingsManager.gErrorsOnly;

            if (SettingsManager.gErrorsOnly)
            {
                toolTip1.SetToolTip(ErrorsOnlyButton, "Show Everything");
                ErrorsOnlyButton.BackColor = SystemColors.ControlDark;
                treeListView1.ExpandAll();
                
                treeListView1.ModelFilter = new ModelFilter((object x) =>
                {
                    if (x is ADREPLDC adREPLDC)
                    {
                        return adREPLDC.DiscoveryIssues;
                    }
                    else if (x is ReplicationNeighbor replicationNeighbor)
                    {
                        return replicationNeighbor.ConsecutiveFailureCount > 0;
                    }

                    return false;
                });
            }
            else
            {
                toolTip1.SetToolTip(ErrorsOnlyButton, "Show Errors Only");

                if (SettingsManager.gDarkMode)
                    ErrorsOnlyButton.BackColor = Color.FromArgb(32, 32, 32);
                else
                    ErrorsOnlyButton.BackColor = SystemColors.Control;

                treeListView1.ModelFilter = null;
            }
        }

        public void OnDiscoveredDCsUpdated()
        {
            DiscoveredDCsUpdated?.Invoke();
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            ProgressPercentLabel.Visible = true;

            ProgressPercentLabel.Text = "0%";

            ActiveForm.Text = $"AD Replication Status Tool - {SettingsManager.ADHelperObject.gForestName}";

            SettingsManager.ADHelperObject.gDCs.Clear();

            foreach (var control in Controls)
            {
                if (control is Button button)
                    button.Enabled = false;
            }

            backgroundWorker1.RunWorkerAsync();
        }

        private async Task RunPing(string destination, AddressFamily addressFamily, Form dialog)
        {
            try
            {
                if (!IPAddress.TryParse(destination, out var address))
                {
                    var entry = await Dns.GetHostEntryAsync(destination);
                    address = entry.AddressList.FirstOrDefault(a => a.AddressFamily == addressFamily);
                    if (address == null)
                    {
                        throw new Exception($"No {addressFamily} address found for {destination}");
                    }
                }

                using (var p = new Ping())
                {
                    var reply = await p.SendPingAsync(address, 5000, new byte[1], new PingOptions(64, true));
                    if (reply.Status == IPStatus.Success)
                    {
                        string protocol = addressFamily == AddressFamily.InterNetwork ? "IPv4" : "IPv6";
                        string successMessage = $"Success:\nDCName: {destination} ({reply.Address})\nProtocol: {protocol}";
                        var statusTextBox = (TextBox)dialog.Controls[2];
                        statusTextBox.Clear();
                        statusTextBox.AppendText($"Ping to {destination} using {protocol} ({reply.Address}) successful.\n");

                        LoggingManager.AppendText($"Ping to {destination} using {protocol} ({reply.Address}) successful.");
                    }                    
                    else
                    {
                        throw new Exception(reply.Status.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                dialog.Invoke(new Action(() =>
                {
                    string errorMessage = $"Ping failed!\n{ex.Message}\n";
                    var statusTextBox = (TextBox)dialog.Controls[2];
                    statusTextBox.Clear();
                    statusTextBox.AppendText($"{errorMessage}\n");
                    LoggingManager.Error(errorMessage);
                }));
            }
        }

        private void SetDarkMode()
        {
            toolTip1.SetToolTip(DarkModeButton, "Light Mode");

            BackColor = Color.FromArgb(32, 32, 32);

            foreach (var control in Controls)
            {
                if (control is Button button)
                {
                    button.BackColor = Color.FromArgb(32, 32, 32);
                }

                if (control is Label label)
                {
                    label.BackColor = Color.FromArgb(32, 32, 32);

                    label.ForeColor = Color.White;
                }
            }

            if (SettingsManager.gLoggingEnabled)
            {
                EnableLoggingButton.BackColor = SystemColors.ControlDark;
            }

            if (SettingsManager.gErrorsOnly)
            {
                ErrorsOnlyButton.BackColor = SystemColors.ControlDark;
            }

            treeListView1.BackColor = Color.FromArgb(32, 32, 32);

            foreach (OLVColumn item in treeListView1.Columns)
            {
                var headerstyle = new HeaderFormatStyle();

                headerstyle.SetBackColor(Color.FromArgb(32, 32, 32));

                headerstyle.SetForeColor(Color.White);

                item.HeaderFormatStyle = headerstyle;
            }
        }

        private void SetDcButton_Click(object sender, EventArgs e)
        {
            SetUserDomainControllerForm setUserDCForm = new SetUserDomainControllerForm();
            LoggingManager.AppendText("SetUserDomainController button was clicked.");

            setUserDCForm.ShowDialog();
            setUserDCForm.Dispose();
        }

        private void SetForestButton_Click(object sender, EventArgs e)
        {
            SetForestNameForm setForestNameForm = new SetForestNameForm();
            LoggingManager.AppendText("SetForestName button was clicked.");

            setForestNameForm.ShowDialog();
            setForestNameForm.Dispose();
        }

        private void SetLightMode()
        {
            toolTip1.SetToolTip(DarkModeButton, "Dark Mode");

            BackColor = SystemColors.Control;

            foreach (var control in Controls)
            {
                if (control is Button button)
                    button.BackColor = SystemColors.Control;

                if (control is Label label)
                {
                    label.BackColor = SystemColors.Control;
                    label.ForeColor = SystemColors.ControlText;
                }
            }

            if (LoggingManager.gLoggingEnabled)
                EnableLoggingButton.BackColor = SystemColors.ControlDark;

            if (SettingsManager.gErrorsOnly)
                ErrorsOnlyButton.BackColor = SystemColors.ControlDark;

            treeListView1.BackColor = SystemColors.Window;

            foreach (OLVColumn item in treeListView1.Columns)
            {
                var headerstyle = new HeaderFormatStyle();

                headerstyle.SetBackColor(SystemColors.Window);

                headerstyle.SetForeColor(SystemColors.ControlText);

                item.HeaderFormatStyle = headerstyle;
            }
        }

        private void SyncDiscoveredDCs()
        {
            SettingsManager.ADHelperObject.gDCs = SettingsManager.ADHelperObject.discoveredDCs.ToList();
            treeListView1.SetObjects(SettingsManager.ADHelperObject.gDCs);
        }

        private void treeListView1_FormatRow(object sender, FormatRowEventArgs e)
        {
            if (e.Model is ADREPLDC dc)
            {
                if (dc.DiscoveryIssues)
                {
                    e.Item.BackColor = Color.Red;
                    e.Item.ForeColor = Color.White;
                }
                else
                {
                    if (SettingsManager.gDarkMode)
                        e.Item.ForeColor = Color.White;
                }
            }
            else if (e.Model is ReplicationNeighbor neighbor)
            {
                if (neighbor.ConsecutiveFailureCount > 0)
                {
                    e.Item.BackColor = Color.Red;
                    e.Item.ForeColor = Color.White;
                }
                else
                {
                    if (SettingsManager.gDarkMode)
                        e.Item.ForeColor = Color.White;
                }
            }
        }
    }
}