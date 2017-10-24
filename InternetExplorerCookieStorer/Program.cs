using IntegrityLevelManagement;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace InternetExplorerCookieStorer
{
    public class Program
    {
        private const int MaximumAllowedIntegrityLevel = NativeMethod.SECURITY_MANDATORY_LOW_RID;

        public string Path
        {
            get
            {
                return (new Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath.Replace("%20", " ");
            }
        }

        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetSetCookieEx(string UrlName, string CookieName, string CookieData, uint dwFlags, IntPtr dwReserved);

        [DllImport("kernel32.dll")]
        public static extern uint GetLastError();

        static void Main(string[] args)
        {
            if (DerivedMethod.GetProcessIntegrityLevel() <= MaximumAllowedIntegrityLevel)
            {
                string cookieHostAndPath = Encoding.UTF8.GetString(Convert.FromBase64String(args[0]));
                string cookieName = Encoding.UTF8.GetString(Convert.FromBase64String(args[1]));
                string cookieValue = Encoding.UTF8.GetString(Convert.FromBase64String(args[2]));
                string expirationDate = new DateTime(long.Parse(args[3])).ToString("R");
                uint cookieFlags = uint.Parse(args[4]);
                if (!InternetSetCookieEx("http://" + cookieHostAndPath, null, cookieName + " = " + cookieValue + "; Expires = " + expirationDate, cookieFlags, IntPtr.Zero))
                {
                    //MessageBox.Show("An error with code " + GetLastError() + " has occurred while storing cookie data");
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
    }
}
