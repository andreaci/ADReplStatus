using BrightIdeasSoftware;
using CredentialManagement;
using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADReplStatus
{
    public partial class ADReplStatusForm : Form
    {
        public static bool gDarkMode;

        public static List<ADREPLDC> gDCs = new List<ADREPLDC>();

        public static bool gErrorsOnly;

        public static string gForestName = string.Empty;

        public static string gLogfileName = string.Empty;

        public static bool gLoggingEnabled;

        public static string gPassword = string.Empty;

        public static string gTarget = string.Empty;

        public static string gUserDomainController = string.Empty;

        public static string gUsername = string.Empty;

        public static bool gUseUserDomainController;

        private readonly ConcurrentBag<ADREPLDC> discoveredDCs = new ConcurrentBag<ADREPLDC>();

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

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\ADREPLSTATUS", false))
                {
                    if (key != null)
                    {
                        gForestName = key.GetValue("ForestName", string.Empty).ToString();

                        gDarkMode = Convert.ToBoolean(key.GetValue("DarkMode", false));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occured while trying to read app settings from the registry!\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (gDarkMode)
            {
                SetDarkMode();
            }
            else
            {
                SetLightMode();
            }

            if (string.IsNullOrEmpty(gForestName))
            {
                try
                {
                    using (Forest forest = Forest.GetCurrentForest())
                    {
                        gForestName = forest.Name;
                    }
                }
                catch
                {
                    MessageBox.Show("Unable to detect AD forest. You will need to manually enter the AD forest you wish to scan using the 'Manually Set Forest' button.\nThis happens on non-domain joined computers as well as hybrid or Azure AD domain-joined machines.", "Forest Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
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

            if (gLoggingEnabled)
            {
                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] AlternateCreds button was clicked.\n");
            }

            alternateCredsForm.ShowDialog();

            alternateCredsForm.Dispose();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Forest forest = null;

            try
            {
                DirectoryContext forestContext;
                var credential = new Credential { Target = "ADCredentials" };
                credential.Load();

                if (credential != null)
                {
                    forestContext = new DirectoryContext(DirectoryContextType.Forest, gForestName, credential.Username, credential.Password);
                }
                else
                {
                    forestContext = new DirectoryContext(DirectoryContextType.Forest, gForestName);
                }

                forest = Forest.GetForest(forestContext);
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(0, $"ERROR: Unable to find AD forest: {gForestName} \n{ex.Message}\n");
                return;
            }

            DomainCollection domainCollection = forest.Domains;
            backgroundWorker1.ReportProgress(0, $"Found {domainCollection.Count} domains in forest {forest.Name}.");

            foreach (Domain domain in domainCollection)
            {
                Parallel.ForEach(domain.DomainControllers.Cast<DomainController>(), dc =>
                {
                    ADREPLDC adrepldc = new ADREPLDC { Name = dc.Name, DomainName = domain.Name };
                    bool discoveryIssues = false;

                    try
                    {
                        using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30)))
                        {
                            adrepldc.Site = dc.SiteName;
                        }
                    }
                    catch (Exception ex)
                    {
                        backgroundWorker1.ReportProgress(0, $"Failed to contact DC {dc.Name} for site name: {ex.Message}");
                        adrepldc.Site = "Unknown";
                        discoveryIssues = true;
                    }

                    try
                    {
                        adrepldc.IsGC = dc.IsGlobalCatalog().ToString();
                    }
                    catch (Exception ex)
                    {
                        backgroundWorker1.ReportProgress(0, $"Failed to determine GC status for {dc.Name}: {ex.Message}");
                        adrepldc.IsGC = "Unknown";
                        discoveryIssues = true;
                    }

                    try
                    {
                        using (DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://" + dc.Name))
                        using (DirectorySearcher search = new DirectorySearcher(directoryEntry))
                        {
                            search.ClientTimeout = TimeSpan.FromSeconds(30);
                            search.Filter = $"(samaccountname={dc.Name.Split('.')[0]}$)";
                            search.PropertiesToLoad.Add("msDS-isRODC");
                            SearchResult result = search.FindOne();
                            adrepldc.IsRODC = result?.Properties.Contains("msDS-isRODC") == true && (bool)result.Properties["msDS-isRODC"][0] ? "True" : "False";
                        }
                    }
                    catch (Exception ex)
                    {
                        backgroundWorker1.ReportProgress(0, $"Failed to determine RODC status for {dc.Name}: {ex.Message}");
                        adrepldc.IsRODC = "Unknown";
                        discoveryIssues = true;
                    }

                    if (!discoveryIssues)
                    {
                        try
                        {
                            foreach (ReplicationNeighbor partner in dc.GetAllReplicationNeighbors())
                            {
                                adrepldc.ReplicationPartners.Add(partner);
                            }
                        }
                        catch (Exception ex)
                        {
                            backgroundWorker1.ReportProgress(0, $"Failed to determine replication neighbors for {dc.Name}: {ex.Message}");
                            discoveryIssues = true;
                        }
                    }

                    adrepldc.DiscoveryIssues = discoveryIssues;
                    discoveredDCs.Add(adrepldc);
                    OnDiscoveredDCsUpdated();
                });
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (gLoggingEnabled)
            {
                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] {e.UserState}\n");
            }

            if (e.UserState.ToString().StartsWith("ERROR:"))
            {
                MessageBox.Show(e.UserState.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.UserState.ToString().Equals("UPDATEPERCENT"))
            {
                ProgressPercentLabel.Text = $"{e.ProgressPercentage}%";
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
            gDarkMode = !gDarkMode;

            if (gDarkMode)
            {
                SetDarkMode();
            }
            else
            {
                SetLightMode();
            }

            try
            {
                var key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\ADREPLSTATUS", true);

                if (key != null)
                {
                    if (gDarkMode)
                    {
                        key.SetValue("DarkMode", 1);
                    }
                    else
                    {
                        key.SetValue("DarkMode", 0);
                    }

                    key.Dispose();
                }
            }
            catch (Exception ex)
            {
                string errorMessage = $"ERROR: Failed to write to the HKCU\\ADREPLSTATUS registry key!\n{ex.Message}\n";

                new Thread(() => MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error));

                if (gLoggingEnabled)
                {
                    File.AppendAllText(gLogfileName, $"[{DateTime.Now}] {errorMessage}\n");
                }
            }
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
                        //Create the menustrip
                        ContextMenuStrip diagnosticMenu = new ContextMenuStrip();

                        //Add menuItem click handler
                        diagnosticMenu.ItemClicked += DiagnosticMenuSelector;

                        //Create a List view of all the diagnostics we want to add
                        ObjectListView olv = e.ListView;

                        //Add the "Ping" option
                        ToolStripMenuItem pingMenuItem = new ToolStripMenuItem(String.Format("Ping"));
                        diagnosticMenu.Items.Add(pingMenuItem);

                        //Add the "RDP" option
                        ToolStripMenuItem rdpMenuItem = new ToolStripMenuItem(String.Format("Initiate RDP connection"));
                        diagnosticMenu.Items.Add(rdpMenuItem);

                        //Add the "Enter-PSSession" option
                        ToolStripMenuItem enterPSSessionMenuItem = new ToolStripMenuItem(String.Format("Enter PowerShell session"));
                        diagnosticMenu.Items.Add(enterPSSessionMenuItem);

                        //Add the "Port Tester" option
                        ToolStripMenuItem portTesterMenuItem = new ToolStripMenuItem(String.Format("Port Tester"));
                        diagnosticMenu.Items.Add(portTesterMenuItem);

                        //Actually attach the menu to the cell
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
                    if (gLoggingEnabled)
                    {
                        File.AppendAllText(gLogfileName, $"[{DateTime.Now}] Diagnostic ping menu opened.\n");
                    }
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

            if (gLoggingEnabled)
            {
                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] Port Tester button was clicked.\n");
            }

            protocolTesterForm.ShowDialog();

            protocolTesterForm.Dispose();
        }

        private void DiagnosticPing(object sender, ToolStripItemClickedEventArgs e)
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
                    if (gDarkMode)
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
                if (gLoggingEnabled)
                {
                    File.AppendAllText(gLogfileName, $"[{DateTime.Now}] Initiating remote powershell session to {treeListView1.SelectedItem.Text}.\n");
                }
                string powershellArgs = $"-NoExit $Cred = Get-Credential;Enter-PSSession -ComputerName {treeListView1.SelectedItem.Text} -Credential $Cred";
                Process.Start("powershell.exe", powershellArgs);
            }
            catch (Exception ex)
            {
                string errorMessage = $"ERROR: Enter-PsSession -ComputerName {treeListView1.SelectedItem.Text} failed!\n{ex.Message}\n";

                new Thread(() => MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error));

                if (gLoggingEnabled)
                {
                    File.AppendAllText(gLogfileName, $"[{DateTime.Now}] {errorMessage}\n");
                }
            }
        }

        private void DiagnosticRdp(object sender, ToolStripItemClickedEventArgs e)
        {
            try
            {
                if (gLoggingEnabled)
                {
                    File.AppendAllText(gLogfileName, $"[{DateTime.Now}] Initiating RDP connection to {treeListView1.SelectedItem.Text}.\n");
                }

                string args = $"/v {treeListView1.SelectedItem.Text}";
                Process.Start("mstsc.exe", args);
            }
            catch (Exception ex)
            {
                string errorMessage = $"ERROR: RDP to {treeListView1.SelectedItem.Text} failed!\n{ex.Message}\n";

                new Thread(() => MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error));

                if (gLoggingEnabled)
                {
                    File.AppendAllText(gLogfileName, $"[{DateTime.Now}] {errorMessage}\n");
                }
            }
        }

        private void EnableLoggingButton_Click(object sender, EventArgs e)
        {
            gLoggingEnabled = !gLoggingEnabled;

            if (gLoggingEnabled)
            {
                toolTip1.SetToolTip(EnableLoggingButton, "Disable Logging");

                EnableLoggingButton.BackColor = SystemColors.ControlDark;

                DateTime Now = DateTime.Now;

                gLogfileName = $"adreplstatus_{Now.Month}.{Now.Day}.{Now.Year}.{Now.Hour}.{Now.Minute}.{Now.Second}.log";

                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] Logging enabled.\n");
            }
            else
            {
                toolTip1.SetToolTip(EnableLoggingButton, "Enable Logging");

                if (gDarkMode)
                {
                    EnableLoggingButton.BackColor = Color.FromArgb(32, 32, 32);
                }
                else
                {
                    EnableLoggingButton.BackColor = SystemColors.Control;
                }

                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] Logging disabled.\n");
            }
        }

        private void ErrorsOnlyButton_Click(object sender, EventArgs e)
        {
            gErrorsOnly = !gErrorsOnly;

            if (gErrorsOnly)
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

                if (gDarkMode)
                {
                    ErrorsOnlyButton.BackColor = Color.FromArgb(32, 32, 32);
                }
                else
                {
                    ErrorsOnlyButton.BackColor = SystemColors.Control;
                }

                treeListView1.ModelFilter = null;
            }
        }

        private void OnDiscoveredDCsUpdated()
        {
            DiscoveredDCsUpdated?.Invoke();
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            ProgressPercentLabel.Visible = true;

            ProgressPercentLabel.Text = "0%";

            ActiveForm.Text = $"AD Replication Status Tool - {gForestName}";

            gDCs.Clear();

            foreach (var control in Controls)
            {
                if (control is Button button)
                {
                    button.Enabled = false;
                }
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

                        if (gLoggingEnabled)
                        {
                            File.AppendAllText(gLogfileName, $"[{DateTime.Now}] {statusTextBox.Text}");
                        }
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
                    if (gLoggingEnabled)
                    {
                        File.AppendAllText(gLogfileName, $"[{DateTime.Now}] {errorMessage}");
                    }
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

            if (gLoggingEnabled)
            {
                EnableLoggingButton.BackColor = SystemColors.ControlDark;
            }

            if (gErrorsOnly)
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

            if (gLoggingEnabled)
            {
                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] SetUserDomainController button was clicked.\n");
            }

            setUserDCForm.ShowDialog();

            setUserDCForm.Dispose();
        }

        private void SetForestButton_Click(object sender, EventArgs e)
        {
            SetForestNameForm setForestNameForm = new SetForestNameForm();

            if (gLoggingEnabled)
            {
                File.AppendAllText(gLogfileName, $"[{DateTime.Now}] SetForestName button was clicked.\n");
            }

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
                {
                    button.BackColor = SystemColors.Control;
                }

                if (control is Label label)
                {
                    label.BackColor = SystemColors.Control;

                    label.ForeColor = SystemColors.ControlText;
                }
            }

            if (gLoggingEnabled)
            {
                EnableLoggingButton.BackColor = SystemColors.ControlDark;
            }

            if (gErrorsOnly)
            {
                ErrorsOnlyButton.BackColor = SystemColors.ControlDark;
            }

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
            gDCs = discoveredDCs.ToList();
            treeListView1.SetObjects(gDCs);
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
                    if (gDarkMode)
                    {
                        e.Item.ForeColor = Color.White;
                    }
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
                    if (gDarkMode)
                    {
                        e.Item.ForeColor = Color.White;
                    }
                }
            }
        }
    }
}