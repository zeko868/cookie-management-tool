using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CookieManagementTool
{
    public partial class ProfileSelection : Form
    {
        private IEnumerable<string> browserProfiles;
        private string initiallySelectedProfile;

        public ProfileSelection(string browserName, IEnumerable<string> browserProfiles, string initiallySelectedProfile = null)
        {
            InitializeComponent();

            this.Text += " (" + browserName + ")";
            this.browserProfiles = browserProfiles;
            this.initiallySelectedProfile = initiallySelectedProfile;
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("You have to select at least one profile if you want to import changes to following web-browser! Press Cancel if you realised that you actually don't want to import them to this one");
            }
            else
            {
                this.Close();
            }
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ProfileSelection_Load(object sender, EventArgs e)
        {
            foreach (string profileName in this.browserProfiles)
            {
                this.dataGridView1.Rows.Add(profileName);
                if (profileName == this.initiallySelectedProfile)
                {
                    this.dataGridView1.ClearSelection();
                    this.dataGridView1.Rows[this.dataGridView1.RowCount - 1].Selected = true;
                }
            }
        }
    }
}
