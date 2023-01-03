using SajjuCode.OutlookAddIns.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SajjuCode.OutlookAddIns
{
    public partial class frmSettings : Form
    {
        public frmSettings()
        {
            InitializeComponent();
        }

        private void cmdMessageSaveIn_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.folderBrowserDialog1.ShowDialog()== DialogResult.OK)
                {
                    this.MsgSaveIntextBox.Text = this.folderBrowserDialog1.SelectedPath;
                }
            }
            catch (Exception ex)
            {

                
            }
        }

        private void cmdBrowseDeals_Click(object sender, EventArgs e)
        {
            try
            {
                this.openFileDialog1.Title = "Offerte CSV";
                this.openFileDialog1.Filter = "CSV|*.csv";
                if (this.openFileDialog1.ShowDialog()== DialogResult.OK)
                {
                    this.OfferCSVtextBox.Text = this.openFileDialog1.FileName;
                }
            }
            catch (Exception ex)
            {

                
            }
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(this.OfferCSVtextBox.Text) && Directory.Exists(this.MsgSaveIntextBox.Text))                    
                {
                    var fileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "nwtDeals.txt");
                    string FileData = this.OfferCSVtextBox.Text + ";" + this.MsgSaveIntextBox.Text + ";";
                    File.WriteAllText(fileName, FileData);
                    SettingsClassFunctions.ReadSettingFile();
                    this.DialogResult = DialogResult.OK;
                }
                else
                {
                    MessageBox.Show("CSV or Message Save In folder not exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {

                
            }
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            try
            {
                if (SettingsClassFunctions.isDealFileExists)
                {
                    this.OfferCSVtextBox.Text = SettingsClassFunctions.OfferCSVPath;
                    this.MsgSaveIntextBox.Text = SettingsClassFunctions.MessageSaveFolder;
                }
            }
            catch (Exception ex)
            {

                
            }
        }
    }
}
