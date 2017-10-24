using System;
using System.Windows.Forms;
using IntegrityLevelManagement;

namespace CookieManagementTool
{
    static class Program
    {
        private const int MinimumRequiredIntegrityLevel = NativeMethod.SECURITY_MANDATORY_MEDIUM_RID;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (DerivedMethod.GetProcessIntegrityLevel() < MinimumRequiredIntegrityLevel)
            {
                try
                {
                    // Try to launch a new instance of the current application at the low 
                    // integrity level.
                    DerivedMethod.CreateSpecificIntegrityProcess(Application.ExecutablePath, MinimumRequiredIntegrityLevel);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "CreateSpecificIntegrityProcess Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainWindow());
            }
        }

    }
}
