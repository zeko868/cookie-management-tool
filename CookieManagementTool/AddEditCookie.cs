using IntegrityLevelManagement;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Windows.Forms;

namespace CookieManagementTool
{
    public partial class AddEditCookie : Form
    {
        private string host;
        private string path;
        private string name;
        private string value;
        private DateTime expirationDate;
        private bool secure;
        private bool httpOnly;
        private bool updating = false;
        private string initiallySelectedBrowser;
        private string initiallySelectedProfile;

        private readonly string[] unsupportedBrowsers = { "edge" };

        private const string allowedChars = "A-Za-z0-9!#$%&'\\(\\)\\*\\+\\-\\.\\/:<>\\?@\\[\\]\\^_`\\{\\|\\}~";

        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetSetCookieEx(string UrlName, string CookieName, string CookieData, uint dwFlags, IntPtr dwReserved);

        [DllImport("kernel32.dll")]
        public static extern uint GetLastError();

        public AddEditCookie(IEnumerable<KeyValuePair<string, string>> webBrowsers, string initiallySelectedBrowser, string initiallySelectedProfile, IEnumerable<string> webDomainList)
        {
            InitializeComponent();
            this.Text = "Add cookie";
            this.comboBox1.DataSource = new BindingSource(webDomainList, null);
            this.initiallySelectedBrowser = initiallySelectedBrowser;
            this.initiallySelectedProfile = initiallySelectedProfile;
            this.dateTimePicker1.Value = DateTime.Now.AddMonths(1);
            int rowNum = 0;
            foreach (KeyValuePair<string, string> pair in webBrowsers)
            {
                if (!this.unsupportedBrowsers.Contains(pair.Key))
                {
                    this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[rowNum].Cells[0].Value = pair.Value;
                    this.dataGridView1.Rows[rowNum++].Cells[0].Tag = pair.Key;
                }
            }
        }

        public AddEditCookie(IEnumerable<KeyValuePair<string, string>> webBrowsers, string initiallySelectedBrowser, string initiallySelectedProfile, IEnumerable<string> webDomainList, string hostAndPath, string name, string value, DateTime expirationDate, bool secure, bool httpOnly)
        {
            InitializeComponent();
            this.Text = "Edit cookie";
            this.updating = true;
            this.initiallySelectedBrowser = initiallySelectedBrowser;
            this.initiallySelectedProfile = initiallySelectedProfile;
            int slashPosition = hostAndPath.IndexOf('/');
            this.comboBox1.DataSource = new BindingSource(webDomainList, null);
            this.host = this.comboBox1.Text = hostAndPath.Substring(0, slashPosition);
            this.path = this.textBox1.Text = hostAndPath.Substring(slashPosition);
            this.name = this.textBox2.Text = name;
            this.value = this.textBox3.Text = value;
            this.expirationDate = this.dateTimePicker1.Value = expirationDate;
            this.checkBox1.Checked = this.secure = secure;
            this.checkBox2.Checked = this.httpOnly = httpOnly;
            int rowNum = 0;
            foreach (KeyValuePair<string, string> pair in webBrowsers)
            {
                if (!this.unsupportedBrowsers.Contains(pair.Key))
                {
                    this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[rowNum].Cells[0].Value = pair.Value;
                    this.dataGridView1.Rows[rowNum++].Cells[0].Tag = pair.Key;
                }
            }
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("At least one web-browser has to be selected");
                return;
            }
            foreach (DataGridViewRow dgvr in this.dataGridView1.SelectedRows)
            {
                string selectedBrowserName = (string)dgvr.Cells[0].Tag;
                if (selectedBrowserName == "edge") {
                    if (this.updating)
                    {
                        if (this.host == this.comboBox1.Text)
                        {
                            string localAppPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                            foreach (string foldername in new string[] { "#!001", "#!002", "" })
                            {
                                string location = Path.Combine(localAppPath, "Packages", "Microsoft.MicrosoftEdge_8wekyb3d8bbwe", "AC", foldername, "MicrosoftEdge", "Cookies");
                                if (this.AddEditCookieOfEdge(location, this.comboBox1.Text + this.textBox1.Text, this.name, this.textBox2.Text, this.textBox3.Text, this.dateTimePicker1.Value.Ticks, this.checkBox1.Checked, this.checkBox2.Checked))
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Editing web-domain information of cookie for Microsoft Edge is still not implemented!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Storing new cookies for Microsoft Edge is still not implemented!");
                    }
                }
                else
                {
                    IEnumerable<string> profileNames;
                    Func<IEnumerable<string>> GetBrowserProfiles;
                    Action<string, bool> StoreCookie;
                    switch (selectedBrowserName)
                    {
                        case "firefox":
                            GetBrowserProfiles = MainWindow.ReadFirefoxProfiles;
                            StoreCookie = this.FirefoxStoreCookie;
                            break;
                        case "chrome":
                            GetBrowserProfiles = MainWindow.ReadChromeProfiles;
                            StoreCookie = this.ChromeStoreCookie;
                            break;
                        default:    // case "explorer":
                            GetBrowserProfiles = MainWindow.ReadInternetExplorerProfiles;
                            StoreCookie = this.InternetExplorerStoreCookie;
                            break;
                    }
                    profileNames = GetBrowserProfiles();
                    if (profileNames.Count() > 1)
                    {
                        ProfileSelection form;
                        if (this.initiallySelectedBrowser == (string)dgvr.Cells[0].Tag)
                        {
                            form = new ProfileSelection((string)dgvr.Cells[0].Value, profileNames, this.initiallySelectedProfile);
                        }
                        else
                        {
                            form = new ProfileSelection((string)dgvr.Cells[0].Value, profileNames);
                        }
                        if (DialogResult.OK == form.ShowDialog())
                        {
                            foreach (DataGridViewRow row in form.dataGridView1.SelectedRows)
                            {
                                StoreCookie((string)row.Cells[0].Value, this.updating);
                            }
                        }
                    }
                    else
                    {
                        StoreCookie(profileNames.First(), this.updating);
                    }
                }
            }
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (this.initiallySelectedBrowser == "all")
            {
                this.dataGridView1.SelectAll();
            }
            else
            {
                this.dataGridView1.ClearSelection();
                foreach (DataGridViewRow dgvr in this.dataGridView1.Rows)
                {
                    if (this.initiallySelectedBrowser == (string)dgvr.Cells[0].Tag)
                    {
                        dgvr.Selected = true;
                    }
                }
            }
        }

        private bool DeleteCookieOfIE(string location, string hostAndPath, string name)
        {
            foreach (string cookieFilename in Directory.GetFiles(location))
            {
                if (cookieFilename.EndsWith(".cookie"))
                {
                    string newFileContent = "";
                    bool fileHasToBeUpdated = false;
                    foreach (string cookieInfo in File.ReadAllText(cookieFilename, Encoding.Default).Split(new string[] { "*\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        string[] cookieInfoSplit = cookieInfo.Split('\n');
                        if (hostAndPath == cookieInfoSplit[2] && name == cookieInfoSplit[0])
                        {
                            fileHasToBeUpdated = true;
                            break;
                        }
                        else
                        {
                            newFileContent += cookieInfo + "*\n";
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
                        return true;
                    }
                }
            }
            return false;
        }

        private bool AddEditCookieOfEdge(string location, string hostAndPath, string nameOld, string nameNew, string valueNew, long expirationTimeNew, bool secureNew,  bool httpOnlyNew)
        {
            foreach (string cookieFilename in Directory.GetFiles(location))
            {
                if (cookieFilename.EndsWith(".cookie"))
                {
                    string newFileContent = "";
                    bool fileHasToBeUpdated = false;
                    foreach (string cookieInfo in File.ReadAllText(cookieFilename, Encoding.Default).Split(new string[] { "*\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        string[] cookieInfoSplit = cookieInfo.Split('\n');
                        if (hostAndPath == cookieInfoSplit[2] && nameOld == cookieInfoSplit[0])
                        {
                            fileHasToBeUpdated = true;
                            uint cookieFlag = 2147484672;
                            if (secureNew)
                            {
                                cookieFlag |= 0x1;
                            }
                            if (httpOnly)
                            {
                                cookieFlag |= 0x2000;
                            }
                            uint first32Bits = (uint)(expirationTimeNew >> 32);
                            uint last32Bits = (uint)expirationTimeNew;
                            newFileContent += String.Join("\n", nameNew, valueNew, hostAndPath, cookieFlag, last32Bits, first32Bits) + "*\n";
                            break;
                        }
                        else
                        {
                            newFileContent += cookieInfo + "*\n";
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
                        return true;
                    }
                }
            }
            return false;
        }

        private void FirefoxStoreCookie(string profileName, bool isUpdating)
        {
            string roamingAppPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string location = Path.Combine(roamingAppPath, "Mozilla", "firefox", "Profiles", profileName, "cookies.sqlite");
            string sqlQuery;

            DateTime startTime = new DateTime(1970, 1, 1);
            const int divisorFrom100NanosecConversion = 10000000;
            long expirationTicks = this.dateTimePicker1.Value.Subtract(startTime).Ticks / divisorFrom100NanosecConversion;
            if (!updating)
            {
                sqlQuery = String.Format("INSERT INTO moz_cookies (host, path, name, value, expiry, isSecure, isHttpOnly) VALUES ('{0}', '{1}', '{2}', '{3}', {4}, {5}, {6})",
                    this.comboBox1.Text,
                    this.textBox1.Text,
                    this.textBox2.Text,
                    this.textBox3.Text,
                    expirationTicks,
                    this.checkBox1.Checked ? 1 : 0,
                    this.checkBox2.Checked ? 1 : 0
                );
            }
            else
            {
                sqlQuery = String.Format("UPDATE moz_cookies SET host='{0}', path='{1}', name='{2}', value='{3}', expiry={4}, isSecure={5}, isHttpOnly={6} WHERE (host='{7}' OR host='.{7}') AND path='{8}' AND name='{9}'",
                    this.comboBox1.Text,
                    this.textBox1.Text,
                    this.textBox2.Text,
                    this.textBox3.Text,
                    expirationTicks,
                    this.checkBox1.Checked ? 1 : 0,
                    this.checkBox2.Checked ? 1 : 0,
                    this.host,
                    this.path,
                    this.name
                );
            }

            SQLiteConnection conn = new SQLiteConnection(String.Format("Data Source={0}", location));
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sqlQuery, conn);
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch
            {
            }
            conn.Close();
        }

        private void ChromeStoreCookie(string profileName, bool isUpdating)
        {
            string localAppPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string location = Path.Combine(localAppPath, "Google", "Chrome", "User Data", profileName, "Cookies");

            DateTime startTime = new DateTime(1601, 1, 1);
            const int divisorFrom100NanosecConversion = 10;
            long expirationTicks = this.dateTimePicker1.Value.Subtract(startTime).Ticks / divisorFrom100NanosecConversion;

            SQLiteConnection conn = new SQLiteConnection(String.Format("Data Source={0}", location));
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(conn);
            bool errorHasOccurred;
            Nullable<bool> previouslyEncrypted = null;
            bool currentlyEncrypted = false;    // set value is for now meaningless
            do
            {
                errorHasOccurred = false;
                if (!updating)
                {
                    if (DialogResult.Yes == MessageBox.Show("Google Chrome supports storage and usage of encrypted cookie values. Do you want to encrypt value of inserted cookie?", String.Format("Cookie value encryption ({0})", profileName), MessageBoxButtons.YesNo))
                    {
                        cmd.CommandText = "INSERT INTO cookies (host_key, path, name, value, encrypted_value, expires_utc, secure, httponly, last_access_utc) VALUES (:host, :path, :name, '', :value, :expiry, :secure, :httponly, 0)";
                        cmd.Parameters.Add("value", DbType.Binary).Value = ProtectedData.Protect(Encoding.UTF8.GetBytes(this.textBox3.Text), null, DataProtectionScope.CurrentUser);
                    }
                    else
                    {
                        cmd.CommandText = "INSERT INTO cookies (host_key, path, name, value, expires_utc, secure, httponly, last_access_utc) VALUES (:host, :path, :name, :value, :expiry, :secure, :httponly, 0)";
                        cmd.Parameters.Add("value", DbType.String).Value = this.textBox3.Text;
                    }
                }
                else
                {
                    if (previouslyEncrypted == null)
                    {
                        cmd.CommandText = "SELECT encrypted_value FROM cookies WHERE (host_key=:hostOld OR host_key=('.' || :hostOld)) AND path=:pathOld AND name=:nameOld";
                        cmd.Parameters.Add("hostOld", DbType.String).Value = this.host;
                        cmd.Parameters.Add("pathOld", DbType.String).Value = this.path;
                        cmd.Parameters.Add("nameOld", DbType.String).Value = this.name;
                        previouslyEncrypted = cmd.ExecuteScalar() != null;
                        currentlyEncrypted = DialogResult.Yes == MessageBox.Show("Until now cookie value " + (previouslyEncrypted.Value ? "was" : "wasn't") + " encrypted. Do you want to keep its current security level?", String.Format("Keep current cookie value's security level (Google Chrome - {0})",  profileName), MessageBoxButtons.YesNo);
                    }
                    if (currentlyEncrypted)
                    {
                        cmd.CommandText = "UPDATE cookies SET host_key=:host, path=:path, name=:name, value='', encrypted_value=:value, expires_utc=:expiry, secure=:secure, httponly=:httponly WHERE (host_key=:hostOld OR host_key=('.' || :hostOld)) AND path=:pathOld AND name=:nameOld";
                        cmd.Parameters.Add("value", DbType.Binary).Value = ProtectedData.Protect(Encoding.UTF8.GetBytes(this.textBox3.Text), null, DataProtectionScope.CurrentUser);
                    }
                    else
                    {
                        cmd.CommandText = "UPDATE cookies SET host_key=:host, path=:path, name=:name, value=:value, encrypted_value=null, expires_utc=:expiry, secure=:secure, httponly=:httponly WHERE (host_key=:hostOld OR host_key=('.' || :hostOld)) AND path=:pathOld AND name=:nameOld";
                        cmd.Parameters.Add("value", DbType.String).Value = this.textBox3.Text;
                    }
                    cmd.Parameters.Add("hostOld", DbType.String).Value = this.host;
                    cmd.Parameters.Add("pathOld", DbType.String).Value = this.path;
                    cmd.Parameters.Add("nameOld", DbType.String).Value = this.name;
                }
                cmd.Parameters.Add("host", DbType.String).Value = this.comboBox1.Text;
                cmd.Parameters.Add("path", DbType.String).Value = this.textBox1.Text;
                cmd.Parameters.Add("name", DbType.String).Value = this.textBox2.Text;
                cmd.Parameters.Add("expiry", DbType.Int64).Value = expirationTicks;
                cmd.Parameters.Add("secure", DbType.Boolean).Value = this.checkBox1.Checked;
                cmd.Parameters.Add("httponly", DbType.Boolean).Value = this.checkBox2.Checked;

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    errorHasOccurred = true;
                }
            } while (errorHasOccurred && DialogResult.Retry == MessageBox.Show("Looks like your browser is currently open what prevents modifying cookie database. If you want to delete their cookies, close your browser and choose Retry button", String.Format("Unable to {0} {1} ({2}) cookie", isUpdating ? "edit" : "add", "Google Chrome", profileName), MessageBoxButtons.RetryCancel));
            conn.Close();
        }

        private void InternetExplorerStoreCookie(string profileName, bool isUpdating)
        {
            uint cookieFlags = 2147484672;
            if (this.checkBox1.Checked)
            {
                cookieFlags |= 0x1;
            }
            if (this.checkBox2.Checked)
            {
                cookieFlags |= 0x2000;
            }

            bool storeForProtectedMode = profileName == "Protected";
            if (isUpdating && !(this.comboBox1.Text == this.host && this.textBox1.Text == this.path && this.textBox2.Text == this.name))
            {
                string localAppPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                this.DeleteCookieOfIE(Path.Combine(localAppPath, "Microsoft", "Windows", "INetCookies", storeForProtectedMode ? "Low" : ""), this.host + "/" + this.path, this.name);
            }
            
            if (storeForProtectedMode)
            {
                try
                {
                    DerivedMethod.CreateSpecificIntegrityProcess(
                        String.Join(" ",
                            new InternetExplorerCookieStorer.Program().Path,
                            Convert.ToBase64String(Encoding.UTF8.GetBytes(this.comboBox1.Text + this.textBox1.Text)),
                            Convert.ToBase64String(Encoding.UTF8.GetBytes(this.textBox2.Text)),
                            Convert.ToBase64String(Encoding.UTF8.GetBytes(this.textBox3.Text)),
                            this.dateTimePicker1.Value.Ticks,
                            cookieFlags
                        ),
                        NativeMethod.SECURITY_MANDATORY_LOW_RID
                    );

                    Thread.Sleep(1000); // the elegant way would be to create a pipe and wait for output which will be produced when process finishes

                }
                catch (Win32Exception e)
                {
                    MessageBox.Show("Following error has occurred while storing Internet Explorer Protected cookie data: " + e.Message + Environment.NewLine + "Make sure that InternerExplorerCookieStorer has been previously deployed");
                }
            }
            else
            {
                if (!InternetSetCookieEx("http://" + this.comboBox1.Text + this.textBox1.Text, null, this.textBox2.Text + " = " + this.textBox3.Text + "; Expires = " + this.dateTimePicker1.Value.ToString("R"), cookieFlags, IntPtr.Zero))
                {
                    MessageBox.Show("An error with code " + GetLastError() + " has occurred while storing Internet Explorer Non-protected cookie data");
                }

                /* another way to store cookie in Internet Explorer
                WebBrowser wb = new WebBrowser();
                wb.Navigate(new Uri("http://" + this.comboBox1.Text + this.textBox1.Text));
                wb.Navigated += Wb_Navigated;
                */
            }
        }

        /*
        private void Wb_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            ((WebBrowser)sender).Document.Cookie = this.textBox2.Text + " = " + this.textBox3.Text + "; Expires = " + this.dateTimePicker1.Value.ToString("R");
        }
        */

        private void textBox2_Leave(object sender, EventArgs e)
        {
            MatchCollection mc = Regex.Matches(this.textBox2.Text, String.Format("[^{0}]", allowedChars));
            if (mc.Count > 0)
            {
                StringBuilder sb = new StringBuilder(this.textBox2.Text);
                int offset = 0;
                foreach (Match match in mc)
                {
                    string encodedChar = HttpUtility.UrlPathEncode(match.Value);
                    sb.Replace(match.Value, encodedChar, match.Index + offset, 1);
                    offset += encodedChar.Length - 1;
                }
                this.textBox2.Text = sb.ToString();
            }
        }
    }
}
