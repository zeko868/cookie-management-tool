using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;

namespace CookieManagementTool
{
    public partial class MainWindow : Form
    {
        private const string webDomainColumnName = "webDomain";
        private const string nameColumnName = "name";
        private const string valueColumnName = "value";
        private const string expirationTimeColumnName = "expirationTime";
        private const string secureColumnName = "secure";
        private const string httpOnlyColumnName = "httpOnly";
        private const string browserNameColumnName = "browserName";

        private Dictionary<string, string> webBrowsers = new Dictionary<string, string>();
        private string filterColumn = "All columns";
        private string filterValue = "";
        private StringComparison filterStringCaseComparisonType = StringComparison.InvariantCultureIgnoreCase;
        private Func<MainWindow, string, string, bool> filterStringExactnessComparer = delegate (MainWindow frm, string s1, string s2)
        {
            return s1.Contains(s2, frm.filterStringCaseComparisonType);
        };

        public MainWindow()
        {
            this.InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.comboBox2.SelectedIndex = 0;
            this.DetectInstalledWebBrowsers();

            this.comboBox1.DisplayMember = "Value";
            this.comboBox1.ValueMember = "Key";
            this.comboBox1.DataSource = new BindingSource(this.webBrowsers, null);
            this.comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;
            this.tabControl1.Selected += TabControl1_Selected;
            if (comboBox1.Items.Count != 0)
            {
                comboBox1.SelectedIndex = 0;
            }
            this.PreviewProfilesAndCookiesOfSelectedBrowser();

            this.comboBox2.SelectedIndexChanged += ComboBox2_SelectedIndexChanged;
        }

        private void DetectInstalledWebBrowsers()
        {
            this.webBrowsers.Clear();
            RegistryKey directoryKey;
            if (Registry.LocalMachine.OpenSubKey("SOFTWARE\\Mozilla\\Firefox\\TaskBarIDs").GetValueNames().Length != 0)
            {
                this.webBrowsers.Add("firefox", "Mozilla Firefox");
            }
            directoryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\WOW6432Node\\Google\\Update");
            if (directoryKey?.GetValue("UninstallCmdLine") != null)
            {
                this.webBrowsers.Add("chrome", "Google Chrome");
            }
            directoryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Internet Explorer\\Main");
            if (directoryKey?.GetValue("x86AppPath") != null)
            {
                this.webBrowsers.Add("explorer", "Internet Explorer");
            }
            directoryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\MicrosoftEdge\\Main");
            if (directoryKey?.GetValue("OperationalData") != null)
            {
                this.webBrowsers.Add("edge", "Microsoft Edge");
            }
            if (this.webBrowsers.Count > 1)
            {
                this.webBrowsers.Add("all", "All web-browsers (all profiles)");
            }
            this.comboBox1.DataSource = new BindingSource(this.webBrowsers, null);
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.tabControl1.TabPages.Clear();
            this.PreviewProfilesAndCookiesOfSelectedBrowser();
        }

        private void PreviewProfilesAndCookiesOfSelectedBrowser()
        {
            this.tabControl1.TabPages.Clear();
            TabPage tab;
            switch (((KeyValuePair<string, string>)this.comboBox1.SelectedItem).Key)
            {
                case "firefox":
                    foreach (string profileName in ReadFirefoxProfiles())
                    {
                        tab = new TabPage();
                        tab.Text = tab.Name = profileName.Substring(profileName.LastIndexOf("/") + 1);
                        this.tabControl1.TabPages.Add(tab);
                    }
                    break;
                case "chrome":
                    foreach (string profileName in ReadChromeProfiles())
                    {
                        tab = new TabPage();
                        tab.Text = tab.Name = profileName;
                        this.tabControl1.TabPages.Add(tab);
                    }
                    break;
                case "explorer":
                    foreach (string profileName in ReadInternetExplorerProfiles())
                    {
                        tab = new TabPage();
                        tab.Text = tab.Name = profileName;
                        this.tabControl1.TabPages.Add(tab);
                    }
                    break;
                case "edge":
                    tab = new TabPage();
                    tab.Text = tab.Name = "Default";
                    this.tabControl1.TabPages.Add(tab);
                    break;
                default:
                    tab = new TabPage();
                    tab.Text = tab.Name = "All profiles";
                    this.tabControl1.TabPages.Add(tab);
                    break;
            }
            this.GetCookiesOfFollowingProfile(this.tabControl1.SelectedTab);
        }

        public static IEnumerable<string> ReadInternetExplorerProfiles()
        {
            return new string[] { "Protected", "Non-protected" };
        }

        public static IEnumerable<string> ReadChromeProfiles()
        {
            string localAppPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string profFileLocation = Path.Combine(localAppPath, "Google", "Chrome", "User Data", "Local State");
            if (File.Exists(profFileLocation))
            {
                JObject joConfiguration = JObject.Parse(File.ReadAllText(profFileLocation, Encoding.Default));
                JObject ojProfiles = (JObject)joConfiguration["profile"]["info_cache"];
                return ojProfiles.Properties().Select(p => p.Name);
            }
            return new List<string>();
        }

        public static IEnumerable<string> ReadFirefoxProfiles()
        {
            string roamingAppPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string mozillaPath = Path.Combine(roamingAppPath, "Mozilla");

            if (System.IO.Directory.Exists(mozillaPath))
            {
                string firefoxPath = Path.Combine(mozillaPath, "firefox");
                if (System.IO.Directory.Exists(firefoxPath))
                {
                    string profileFile = Path.Combine(firefoxPath, "profiles.ini");

                    if (File.Exists(profileFile))
                    {
                        StreamReader rdr = new StreamReader(profileFile);
                        string resp = rdr.ReadToEnd();
                        string[] lines = resp.Split(new string[] { "\r\n" }, StringSplitOptions.None);

                        return lines.Where(x => x.Contains("Path=")).Select(x => x.Substring(x.IndexOf("/") + 1));
                    }
                }
            }
            return new List<string>();
        }

        private void TabControl1_Selected(object sender, TabControlEventArgs e)
        {
            TabPage tab = e.TabPage;
            if (tab != null)
            {
                tab.Controls.Clear();
                this.GetCookiesOfFollowingProfile(tab);
            }
        }

        private void FillDgvWithCookiesOfFollowingProfile(DataGridView dgv, string selectedBrowser, string profileName)
        {
            string location;
            string selectedBrowserFullname = this.webBrowsers[selectedBrowser];
            string localAppPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string roamingAppPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            if (selectedBrowser == "explorer")
            {
                location = Path.Combine(localAppPath, "Microsoft", "Windows", "INetCookies", profileName == "Protected" ? "Low" : "");
                this.DisplayCookiesOfIE(dgv, location, "explorer", profileName);
            }
            else if (selectedBrowser == "edge")
            {
                foreach (string foldername in new string[] { "#!001", "#!002", "" })
                {
                    location = Path.Combine(localAppPath, "Packages", "Microsoft.MicrosoftEdge_8wekyb3d8bbwe", "AC", foldername, "MicrosoftEdge", "Cookies");
                    this.DisplayCookiesOfIE(dgv, location, "edge");
                }
            }
            else
            {
                if (selectedBrowser == "firefox")
                {
                    location = Path.Combine(roamingAppPath, "Mozilla", "firefox", "Profiles", profileName, "cookies.sqlite");
                }
                else
                {
                    location = Path.Combine(localAppPath, "Google", "Chrome", "User Data", profileName, "Cookies");
                }
                SQLiteConnection conn = new SQLiteConnection(String.Format("Data Source={0}", location));

                conn.Open();

                string sqlQuery = null;
                DateTime startTime = new DateTime();
                int multiplierFor100NanosecConversion = 10000000;
                switch (selectedBrowser)
                {
                    case "firefox":
                        //sqlQuery = "SELECT datetime(expiry + (strftime('%s', '1970-01-01')), 'unixepoch'), host || path, name, value FROM moz_cookies";
                        sqlQuery = "SELECT isSecure, isHttpOnly, expiry, host || path AS hostandpath, name, value FROM moz_cookies";
                        /*
                                                if (filterValue != String.Empty)
                                                {
                                                    switch (filterColumn)
                                                    {
                                                        case "Web-domain":
                                                            sqlQuery += " WHERE hostandpath LIKE '%" + filterValue + "%'";
                                                            break;
                                                        case "Name":
                                                            sqlQuery += " WHERE name LIKE '%" + filterValue + "%'";
                                                            break;
                                                        case "Value":
                                                            sqlQuery += " WHERE value LIKE '%" + filterValue + "%'";
                                                            break;
                                                        default:
                                                            sqlQuery += " WHERE hostandpath LIKE '%" + filterValue + "%' OR name LIKE '%" + filterValue + "%' OR value LIKE '%" + filterValue + "%'";
                                                            break;
                                                    }
                                                }*/
                        startTime = new DateTime(1970, 1, 1);
                        break;
                    case "chrome":
                        //sqlQuery = "SELECT datetime(expires_utc / 1000000 + (strftime('%s', '1601-01-01')), 'unixepoch'), host_key || path, name, value, encrypted_value FROM cookies";
                        sqlQuery = "SELECT secure, httponly, expires_utc, host_key || path, name, value, encrypted_value FROM cookies";
                        startTime = new DateTime(1601, 1, 1);
                        multiplierFor100NanosecConversion = 10;
                        break;
                }
                SQLiteCommand cmd = new SQLiteCommand(sqlQuery, conn);
                //SQLiteCommand cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY 1", conn);
                SQLiteDataReader dr = null;
                bool errorHasOccurred;
                do {
                    try
                    {
                        dr = cmd.ExecuteReader();

                        errorHasOccurred = false;
                    }
                    catch
                    {
                        errorHasOccurred = true;
                        //usually no error occurs while trying to read cookies even if web-browser is open
                        if (DialogResult.Retry != MessageBox.Show("Looks like your browser is currently open what prevents reading cookie database. If you want to read their cookies, close your browser and choose Retry button", String.Format("Unable to read {0} ({1}) cookies", this.webBrowsers[selectedBrowser], profileName), MessageBoxButtons.RetryCancel))
                        {
                            conn.Close();
                            return;
                        }
                    }
                } while (errorHasOccurred);
                while (dr.Read())
                {
                    bool cookieMatchesFilter;
                    string cookieValue = (string)dr.GetValue(5);
                    string cookieWebsite = (string)dr.GetValue(3);
                    if (cookieWebsite[0] == '.')
                    {
                        cookieWebsite = cookieWebsite.Substring(1);
                    }
                    if (selectedBrowser == "chrome")
                    {
                        if (cookieValue == String.Empty)
                        {
                            byte[] decodedData = ProtectedData.Unprotect((byte[])dr.GetValue(6), null, DataProtectionScope.CurrentUser);
                            string plainText = Encoding.UTF8.GetString(decodedData);

                            cookieValue = plainText;
                        }
                    }
                    if (this.filterValue == String.Empty)
                    {
                        cookieMatchesFilter = true;
                    }
                    else
                    {
                        cookieMatchesFilter = this.CheckIfCookieMatchesFilter(cookieWebsite, (string)dr.GetValue(4), cookieValue);
                    }
                    if (cookieMatchesFilter)
                    {
                        DateTime expires;
                        try
                        {
                            expires = startTime.Add(new TimeSpan(dr.GetInt64(2) * multiplierFor100NanosecConversion));
                        }
                        catch
                        {
                            expires = DateTime.MaxValue;
                        }
                        if (dgv.Columns[browserNameColumnName] == null)
                        {
                            dgv.Rows.Add(
                                cookieWebsite,
                                (string)dr.GetValue(4),
                                cookieValue,
                                //dr.GetDateTime(2).ToString(),
                                expires,
                                dr.GetBoolean(0),
                                dr.GetBoolean(1)
                            );
                        }
                        else
                        {
                            dgv.Rows.Add(
                                selectedBrowserFullname + " (" + profileName + ")",
                                cookieWebsite,
                                (string)dr.GetValue(4),
                                cookieValue,
                                //dr.GetDateTime(2).ToString(),
                                expires,
                                dr.GetBoolean(0),
                                dr.GetBoolean(1)
                            );
                            dgv[browserNameColumnName, dgv.RowCount - 1].Tag = new Dictionary<string, string>()
                            {
                                { "browser", selectedBrowser },
                                { "profile", profileName }
                            };
                        }
                    }
                }
                conn.Close();
            }
        }

        private void GetCookiesOfFollowingProfile(TabPage tab)
        {
            DataGridView dgv = new DataGridView();
            dgv.Columns.AddRange(
                new DataGridViewTextBoxColumn()
                {
                    HeaderText = "web-domain",
                    Name = webDomainColumnName,
                    ReadOnly = true
                },
                new DataGridViewTextBoxColumn()
                {
                    HeaderText = "name",
                    Name = nameColumnName,
                    ReadOnly = true
                },
                new DataGridViewTextBoxColumn()
                {
                    HeaderText = "value",
                    Name = valueColumnName,
                    ReadOnly = true
                },
                new DataGridViewTextBoxColumn()
                {
                    HeaderText = "expiration time",
                    Name = expirationTimeColumnName,
                    ReadOnly = true
                },
                new DataGridViewTextBoxColumn()
                {
                    HeaderText = "secure",
                    Name = secureColumnName,
                    ReadOnly = true
                },
                new DataGridViewTextBoxColumn()
                {
                    HeaderText = "http-only",
                    Name = httpOnlyColumnName,
                    ReadOnly = true
                }
            );
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowUserToOrderColumns = true;
            dgv.AllowUserToResizeRows = false;
            dgv.ReadOnly = true;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.Width = tab.Width;
            dgv.Height = tab.Height;
            dgv.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            dgv.Columns[expirationTimeColumnName].ValueType = typeof(DateTime);
            dgv.Columns[secureColumnName].ValueType = dgv.Columns[httpOnlyColumnName].ValueType = typeof(bool);
            dgv.KeyDown += Dgv_KeyDown;

            string selectedBrowser = ((KeyValuePair<string, string>)this.comboBox1.SelectedItem).Key;
            if (selectedBrowser == "all")
            {
                dgv.Columns.Insert(0, new DataGridViewTextBoxColumn()
                {
                    HeaderText = "browser name",
                    Name = browserNameColumnName,
                    ReadOnly = true,
                });
                foreach (string browserName in this.webBrowsers.Keys.Except(new string[] { "all" }))
                {
                    IEnumerable<string> profileNames;
                    switch (browserName)
                    {
                        case "firefox":
                            profileNames = ReadFirefoxProfiles();
                            break;
                        case "chrome":
                            profileNames = ReadChromeProfiles();
                            break;
                        case "explorer":
                            profileNames = ReadInternetExplorerProfiles();
                            break;
                        default:
                            profileNames = new string[] { "Default" };
                            break;
                    }
                    foreach (string profileName in profileNames)
                    {
                        this.FillDgvWithCookiesOfFollowingProfile(dgv, browserName, profileName);
                    }
                }
            }
            else
            {
                this.FillDgvWithCookiesOfFollowingProfile(dgv, selectedBrowser, tab.Text);
            }

            tab.Controls.Clear();
            tab.Controls.Add(dgv);
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if (e.KeyCode == Keys.Home)
            {
                e.Handled = true;
                dgv.FirstDisplayedScrollingRowIndex = 0;
                if (e.Shift)
                {
                    if (!e.Control)
                    {
                        dgv.ClearSelection();
                    }
                    for (int i=0; i < dgv.CurrentCell.RowIndex; i++)
                    {
                        dgv.Rows[i].Selected = true;
                    }
                }
                else
                {
                    dgv.ClearSelection();
                    dgv.Rows[0].Selected = true;
                    dgv.CurrentCell = dgv[0, 0];
                }
            }
            else if (e.KeyCode == Keys.End)
            {
                e.Handled = true;
                dgv.FirstDisplayedScrollingRowIndex = dgv.RowCount - 1;
                if (e.Shift)
                {
                    if (!e.Control)
                    {
                        dgv.ClearSelection();
                    }
                    for (int i = dgv.CurrentCell.RowIndex; i < dgv.RowCount; i++)
                    {
                        dgv.Rows[i].Selected = true;
                    }
                }
                else
                {
                    dgv.ClearSelection();
                    dgv.Rows[dgv.RowCount - 1].Selected = true;
                    dgv.CurrentCell = dgv[0, dgv.RowCount - 1];
                }
            }
        }

        private bool CheckIfCookieMatchesFilter(string hostAndPath, string name, string value)
        {
            switch (this.filterColumn)
            {
                case "Web-domain":
                    if (this.filterStringExactnessComparer(this, hostAndPath, this.filterValue))
                    {
                        return true;
                    }
                    break;
                case "Name":
                    if (this.filterStringExactnessComparer(this, name, this.filterValue))
                    {
                        return true;
                    }
                    break;
                case "Value":
                    if (this.filterStringExactnessComparer(this, value, this.filterValue))
                    {
                        return true;
                    }
                    break;
                default:
                    if (this.filterStringExactnessComparer(this, hostAndPath, this.filterValue) || this.filterStringExactnessComparer(this, name, this.filterValue) || this.filterStringExactnessComparer(this, value, this.filterValue))
                    {
                        return true;
                    }
                    break;
            }
            return false;
        }

        public void DisplayCookiesOfIE(DataGridView dgv, string location, string browserName, string profileName = "Default")
        {
            foreach (string cookieFilename in Directory.GetFiles(location))
            {
                if (cookieFilename.EndsWith(".cookie"))
                {
                    foreach (string cookieInfo in File.ReadAllText(cookieFilename, Encoding.Default).Split(new string[] { "*\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        string[] cookieInfoSplit = cookieInfo.Split('\n');

                        long numOf100NanosecFromWin32Epoch = uint.Parse(cookieInfoSplit[5]);    //number of 100 nanoseconds past from 1st January 1601 represented as 64-bit integer
                        numOf100NanosecFromWin32Epoch <<= 32;
                        numOf100NanosecFromWin32Epoch += uint.Parse(cookieInfoSplit[4]);
                        uint cookieFlags = uint.Parse(cookieInfoSplit[3]);
                        bool secureFlag = (cookieFlags & 0x1) != 0;
                        bool httpOnlyFlag = (cookieFlags & 0x2000) != 0;
                        bool cookieMatchesFilter;
                        if (this.filterValue == String.Empty)
                        {
                            cookieMatchesFilter = true;
                        }
                        else
                        {
                            cookieMatchesFilter = this.CheckIfCookieMatchesFilter(cookieInfoSplit[2], cookieInfoSplit[0], cookieInfoSplit[1]);
                        }
                        if (cookieMatchesFilter) {
                            if (dgv.Columns[browserNameColumnName] == null)
                            {
                                dgv.Rows.Add(
                                    cookieInfoSplit[2],
                                    cookieInfoSplit[0],
                                    cookieInfoSplit[1],
                                    DateTime.FromFileTime(numOf100NanosecFromWin32Epoch),
                                    secureFlag,
                                    httpOnlyFlag
                                );
                            }
                            else
                            {
                                dgv.Rows.Add(
                                    this.webBrowsers[browserName] + " (" + profileName + ")",
                                    cookieInfoSplit[2],
                                    cookieInfoSplit[0],
                                    cookieInfoSplit[1],
                                    DateTime.FromFileTime(numOf100NanosecFromWin32Epoch),
                                    secureFlag,
                                    httpOnlyFlag
                                );
                                dgv[browserNameColumnName, dgv.RowCount - 1].Tag = new Dictionary<string, string>()
                                {
                                    { "browser", browserName },
                                    { "profile", profileName }
                                };
                            }
                        }
                    }
                }
            }
        }

        private void refreshBtn_Click(object sender, EventArgs e)
        {
            string previousBrowser = ((KeyValuePair<string, string>)this.comboBox1.SelectedItem).Key;
            string previousProfile = this.tabControl1.SelectedTab.Name;
            this.DetectInstalledWebBrowsers();
            string currentBrowser = ((KeyValuePair<string, string>)comboBox1.SelectedItem).Key;
            if (currentBrowser != previousBrowser && this.webBrowsers.ContainsKey(previousBrowser))
            {
                this.comboBox1.SelectedValue = previousBrowser;
            }
            else
            {
                this.PreviewProfilesAndCookiesOfSelectedBrowser();
            }
            if (this.tabControl1.SelectedTab.Name != previousProfile && this.tabControl1.TabPages.ContainsKey(previousProfile))
            {
                this.tabControl1.SelectTab(previousProfile);
            }
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            DataGridView dgv = (DataGridView)this.tabControl1.SelectedTab.Controls[0];
            IEnumerable<string> webDomainList = dgv.RowCount > 0 ? new SortedSet<string>(dgv.Rows.Cast<DataGridViewRow>().Select(x => Regex.Match((string)x.Cells[webDomainColumnName].Value, "^[^/]*").Value)) : null;
            if (DialogResult.OK == new AddEditCookie(this.webBrowsers.Where(x => x.Key != "all"), ((KeyValuePair<string, string>)this.comboBox1.SelectedItem).Key, this.tabControl1.SelectedTab.Text, webDomainList).ShowDialog())
            {
                this.refreshBtn.PerformClick();
            }

        }

        private void editBtn_Click(object sender, EventArgs e)
        {
            DataGridView dgv = (DataGridView)this.tabControl1.SelectedTab.Controls[0];
            DataGridViewSelectedRowCollection dgvsrc = dgv.SelectedRows;
            switch (dgvsrc.Count)
            {
                case 0:
                    MessageBox.Show("You have first to select cookie that you want to edit");
                    break;
                case 1:
                    DataGridViewCellCollection dgvcc = dgvsrc[0].Cells;
                    string browserName = ((KeyValuePair<string, string>)this.comboBox1.SelectedItem).Key;
                    string profileName;
                    if (browserName == "all")
                    {
                        Dictionary<string, string> browserAndProfileInfo = (Dictionary<string, string>)dgvcc[browserNameColumnName].Tag;
                        browserName = browserAndProfileInfo["browser"];
                        profileName = browserAndProfileInfo["profile"];
                    }
                    else
                    {
                        profileName = this.tabControl1.SelectedTab.Text;
                    }
                    IEnumerable<string> webDomainList = dgv.RowCount > 0 ? new SortedSet<string>(dgv.Rows.Cast<DataGridViewRow>().Select(x => Regex.Match((string)x.Cells[webDomainColumnName].Value, "^[^/]*").Value)) : null;
                    if (DialogResult.OK == new AddEditCookie(this.webBrowsers.Where(x => x.Key != "all"), browserName, profileName, webDomainList, (string)dgvcc[webDomainColumnName].Value, (string)dgvcc[nameColumnName].Value, (string)dgvcc[valueColumnName].Value, (DateTime)dgvcc[expirationTimeColumnName].Value, (bool)dgvcc[secureColumnName].Value, (bool)dgvcc[httpOnlyColumnName].Value).ShowDialog())
                    {
                        this.refreshBtn.PerformClick();
                    }
                    break;
                default:
                    MessageBox.Show("At the same time can be only one cookie edited");
                    break;
            }
        }

        private void deleteBtn_Click(object sender, EventArgs e)
        {
            DataGridView dgv = (DataGridView)this.tabControl1.SelectedTab.Controls[0];
            DataGridViewSelectedRowCollection dgvsrc = dgv.SelectedRows;
            if (dgvsrc.Count == 0)
            {
                MessageBox.Show("You have first to select cookie that you want to delete");
            }
            else
            {
                string browserName = ((KeyValuePair<string, string>)this.comboBox1.SelectedItem).Key;
                string profileName = this.tabControl1.SelectedTab.Text;
                if (browserName == "all")
                {
                    int browserNameColumnIndex = dgv.Columns[browserNameColumnName].Index;
                    foreach (string eachBrowserName in this.webBrowsers.Keys)
                    {
                        if (eachBrowserName != "all")
                        {
                            IEnumerable<DataGridViewRow> rowsWithCookiesOfCurrentBrowserToDelete = dgvsrc.Cast<DataGridViewRow>().Where(x => ((Dictionary<string, string>)x.Cells[browserNameColumnIndex].Tag)["browser"] == eachBrowserName);
                            foreach (var rowsWithCookiesOfCurrentProfileToDelete in rowsWithCookiesOfCurrentBrowserToDelete.GroupBy(x => ((Dictionary<string, string>)x.Cells[browserNameColumnIndex].Tag)["profile"])) {
                                this.DeleteRowsWithCookiesOfSpecificBrowser(eachBrowserName, rowsWithCookiesOfCurrentProfileToDelete.Key, dgv, rowsWithCookiesOfCurrentProfileToDelete);
                            }
                        }
                    }
                }
                else
                {
                    this.DeleteRowsWithCookiesOfSpecificBrowser(browserName, profileName, dgv, dgvsrc.Cast<DataGridViewRow>());
                }
            }
        }

        private void DeleteRowsWithCookiesOfSpecificBrowser(string browserName, string profileName, DataGridView dgv, IEnumerable<DataGridViewRow> dgvsrc)
        {
            string localAppPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (browserName == "edge" || browserName == "explorer")
            {
                if (browserName == "edge")
                {
                    foreach (string foldername in new string[] { "#!001", "#!002", "" })
                    {
                        string location = Path.Combine(localAppPath, "Packages", "Microsoft.MicrosoftEdge_8wekyb3d8bbwe", "AC", foldername, "MicrosoftEdge", "Cookies");
                        this.DeleteCookiesOfIE(location, dgv, dgvsrc);
                    }
                }
                else
                {
                    this.DeleteCookiesOfIE(Path.Combine(localAppPath, "Microsoft", "Windows", "INetCookies", profileName == "Protected" ? "Low" : ""), dgv, dgvsrc);
                }
                foreach (DataGridViewRow row in dgvsrc)
                {
                    dgv.Rows.Remove(row);
                }
            }
            else
            {
                string roamingAppPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string location;
                if (browserName == "firefox")
                {
                    location = Path.Combine(roamingAppPath, "Mozilla", "firefox", "Profiles", profileName, "cookies.sqlite");
                }
                else
                {
                    location = Path.Combine(localAppPath, "Google", "Chrome", "User Data", profileName, "Cookies");
                }
                SQLiteConnection conn = new SQLiteConnection(String.Format("Data Source={0}", location));
                string sqlQuery;
                string hostColumnName;
                if (browserName == "firefox")
                {
                    sqlQuery = "DELETE FROM moz_cookies";
                    hostColumnName = "host";
                }
                else
                {
                    sqlQuery = "DELETE FROM cookies";
                    hostColumnName = "host_key";
                }
                bool firstCookieToDelete = true;
                foreach (DataGridViewRow row in dgvsrc)
                {
                    string hostAndPath = (string)row.Cells[webDomainColumnName].Value;
                    int slashPosition = (hostAndPath).IndexOf('/');
                    sqlQuery += String.Format(" {0} ({1}='{2}' OR {1}='.{2}') AND path='{3}' AND name='{4}'",
                        firstCookieToDelete ? "WHERE" : "OR",
                        hostColumnName,
                        hostAndPath.Substring(0, slashPosition),
                        hostAndPath.Substring(slashPosition),
                        (string)row.Cells[nameColumnName].Value
                    );
                    firstCookieToDelete = false;
                }
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand(sqlQuery, conn);
                bool errorHasOccurred = false;
                try
                {
                    cmd.ExecuteNonQuery();

                    foreach (DataGridViewRow row in dgvsrc)
                    {
                        dgv.Rows.Remove(row);
                    }
                }
                catch
                {
                    errorHasOccurred = true;
                }
                conn.Close();
                if (errorHasOccurred)
                {
                    if (DialogResult.Retry == MessageBox.Show("Looks like your browser is currently open what prevents modifying cookie database. If you want to delete their cookies, close your browser and choose Retry button", String.Format("Unable to delete {0} ({1}) cookies", this.webBrowsers[browserName], profileName), MessageBoxButtons.RetryCancel))
                    {
                        this.DeleteRowsWithCookiesOfSpecificBrowser(browserName, profileName, dgv, dgvsrc);
                    }
                }
            }
        }

        private void textBox1_DelayedTextChanged(object sender, EventArgs e)
        {
            this.filterValue = this.textBox1.Text;
            this.GetCookiesOfFollowingProfile(this.tabControl1.SelectedTab);
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.filterColumn = this.comboBox2.Text;
            this.GetCookiesOfFollowingProfile(this.tabControl1.SelectedTab);
        }

        private void DeleteCookiesOfIE(string location, DataGridView dgv, IEnumerable<DataGridViewRow> dgvsrc)
        {
            foreach (string cookieFilename in Directory.GetFiles(location))
            {
                if (cookieFilename.EndsWith(".cookie"))
                {
                    string newFileContent = "";
                    bool fileHasToBeUpdated = false;
                    foreach (string cookieInfo in File.ReadAllText(cookieFilename, Encoding.Default).Split(new string[] { "*\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        bool cookieHasToBeDeleted = false;
                        string[] cookieInfoSplit = cookieInfo.Split('\n');
                        foreach (DataGridViewRow row in dgvsrc)
                        {
                            if ((string)row.Cells[webDomainColumnName].Value == cookieInfoSplit[2] && (string)row.Cells[nameColumnName].Value == cookieInfoSplit[0])
                            {
                                cookieHasToBeDeleted = true;
                                row.Selected = false;
                                break;
                            }
                        }
                        if (!cookieHasToBeDeleted)
                        {
                            newFileContent += cookieInfo + "*\n";
                        }
                        else
                        {
                            fileHasToBeUpdated = true;
                        }
                    }
                    if (fileHasToBeUpdated)
                    {
                        if (newFileContent == String.Empty)
                        {
                            File.Delete(cookieFilename);
                        }
                        else
                        {
                            File.WriteAllText(cookieFilename, newFileContent, Encoding.Default);
                        }
                    }
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.filterStringCaseComparisonType = this.checkBox1.Checked ? StringComparison.InvariantCulture : StringComparison.InvariantCultureIgnoreCase;
            this.GetCookiesOfFollowingProfile(this.tabControl1.SelectedTab);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox2.Checked)
            {
                filterStringExactnessComparer = delegate (MainWindow frm, string s1, string s2)
                {
                    return s1.Equals(s2, frm.filterStringCaseComparisonType);
                };
            }
            else
            {
                filterStringExactnessComparer = delegate (MainWindow frm, string s1, string s2)
                {
                    return s1.Contains(s2, frm.filterStringCaseComparisonType);
                };
            }
            this.GetCookiesOfFollowingProfile(this.tabControl1.SelectedTab);
        }
    }

    public static class StringExtensions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) != -1;
        }
    }
}
