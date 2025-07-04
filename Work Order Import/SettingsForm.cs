using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Work_Order_Import
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {

        }
        private void LoadSettings()
        {
            comboBox1.Items.AddRange(new string[] { "Eugene", "Seneca", "Springfield", "Houston", "Demo" });
            comboBox1.Text = Properties.Settings.Default.nestingLocation;
            textBox1.Text = Properties.Settings.Default.demoDbPath;
            textBoxFileLocation.Text = Properties.Settings.Default.fileLocation;
        }

        private void buttonBrowse_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "SQLite DB (*.db)|*.db";
            openFileDialog.Title = "Select Demo SQLite Database";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
            }
        }

        private void buttonSave_Click_1(object sender, EventArgs e)
        {
            Properties.Settings.Default.nestingLocation = comboBox1.Text;
            Properties.Settings.Default.demoDbPath = textBox1.Text;
            Properties.Settings.Default.Save();
            Properties.Settings.Default.fileLocation = textBoxFileLocation.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("Settings saved.");
            this.Close();
        }

        private void buttonCancel_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonFileLocation_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select folder to store part matching file (User Settings.xml)";
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    textBoxFileLocation.Text = folderDialog.SelectedPath;
                }
            }
        }
    }
}
