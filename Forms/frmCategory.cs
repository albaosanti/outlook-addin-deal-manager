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
    public partial class frmCategory : Form
    {
        DataTable myDataTable = new DataTable();
        public string FilePath = SettingsClassFunctions.OfferCSVPath; //"C:\\Deals\\deal list.csv";
        public string FolderPath = SettingsClassFunctions.OfferCSVFolder;// "C:\\Deals";

        void setDataTableColumns()
        {
            try
            {
                myDataTable.Columns.Add(new DataColumn()
                {
                    ColumnName="Number"
                });
                myDataTable.Columns.Add(new DataColumn()
                {
                    ColumnName = "Name"
                });
            }
            catch (Exception ex)
            {

                
            }
        }

        public void validateFileFolder()
        {
            try
            {
              if (!Directory.Exists(FolderPath))
                  Directory.CreateDirectory(FolderPath);

              if (!File.Exists(FilePath))
                File.WriteAllText(FilePath, "Index;Deal name");
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }
    
        void getDeals()
        {
            try
            {
                validateFileFolder();

                if (File.Exists(FilePath))
                {
                    try
                    {
                        int LineNo = 0;
                        string DealNo = "";
                        string DealText = "";
                        foreach (string myDealLine in File.ReadAllLines(FilePath))
                        {
                            if (LineNo == 0)
                            {
                                LineNo++;
                                continue;
                            }

                            DealNo = "";
                            DealText = "";
                            foreach (string myDealInfo in myDealLine.Split(';'))
                            {
                                if (string.IsNullOrEmpty(DealNo))
                                {
                                    DealNo = myDealInfo;
                                }
                                else
                                {
                                    DealText = myDealInfo;
                                }
                            }
                            DataRow myRow = myDataTable.NewRow();
                            myRow["Number"] = DealNo;
                            myRow["Name"] = DealText;
                            myDataTable.Rows.Add(myRow);
                        }
                    }
                    catch (Exception ex)
                    {
                      Console.WriteLine(ex.Message);
                    }

                }

                this.MaindataGridView.AutoGenerateColumns = false;
                if (myDataTable !=null)
                {
                    MaindataGridView.DataSource = null;
                    this.clmDealName.DataPropertyName = "Name";
                    this.clmNO.DataPropertyName = "Number";
                    this.MaindataGridView.DataSource = myDataTable;
                    this.MaindataGridView.Refresh();
                }

            }
            catch (Exception ex)
            {
            }
        }

        public void updateDeals()
        {
            try
            {
                if (MessageBox.Show("Vuoi aggiornare le offerte?","Confirmation",MessageBoxButtons.YesNo,MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    return;
                }

                StringBuilder xmlString = new StringBuilder("Index;Deal name");
                foreach(DataGridViewRow myGridRow in this.MaindataGridView.Rows)
                {
                    try
                    {
                        if (myGridRow.Cells[clmNO.Name].Value !=null && myGridRow.Cells[clmDealName.Name].Value !=null)
                        {
                            xmlString.AppendLine(myGridRow.Cells[clmNO.Name].Value.ToString() + ";" + myGridRow.Cells[clmDealName.Name].Value.ToString());
                        }
                    }
                    catch (Exception ex)
                    {   
                    }
                }

                File.WriteAllText(FilePath, xmlString.ToString());

                MessageBox.Show("Data Saved Successfully.");
                this.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        public frmCategory()
        {
            InitializeComponent();
        }

        private void frmCategory_Load(object sender, EventArgs e)
        {
            try
            {
                this.setDataTableColumns();
                this.MaindataGridView.MultiSelect = false;
                this.getDeals();
            }
            catch (Exception ex)
            {
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            this.updateDeals();
        }
    }
}
