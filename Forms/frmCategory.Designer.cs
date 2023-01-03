
namespace SajjuCode.OutlookAddIns
{
    partial class frmCategory
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlFooter = new System.Windows.Forms.Panel();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdSave = new System.Windows.Forms.Button();
            this.MaindataGridView = new System.Windows.Forms.DataGridView();
            this.clmNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmDealName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnlFooter.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MaindataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlFooter
            // 
            this.pnlFooter.Controls.Add(this.cmdCancel);
            this.pnlFooter.Controls.Add(this.cmdSave);
            this.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlFooter.Location = new System.Drawing.Point(0, 530);
            this.pnlFooter.Name = "pnlFooter";
            this.pnlFooter.Size = new System.Drawing.Size(699, 59);
            this.pnlFooter.TabIndex = 0;
            // 
            // cmdCancel
            // 
            this.cmdCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Location = new System.Drawing.Point(579, 6);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(117, 41);
            this.cmdCancel.TabIndex = 1;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdSave
            // 
            this.cmdSave.Enabled = false;
            this.cmdSave.Location = new System.Drawing.Point(438, 6);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(117, 41);
            this.cmdSave.TabIndex = 0;
            this.cmdSave.Text = "Save";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Visible = false;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // MaindataGridView
            // 
            this.MaindataGridView.AllowUserToOrderColumns = true;
            this.MaindataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.MaindataGridView.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            this.MaindataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MaindataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clmNO,
            this.clmDealName});
            this.MaindataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MaindataGridView.Location = new System.Drawing.Point(0, 0);
            this.MaindataGridView.Name = "MaindataGridView";
            this.MaindataGridView.Size = new System.Drawing.Size(699, 530);
            this.MaindataGridView.TabIndex = 1;
            // 
            // clmNO
            // 
            this.clmNO.HeaderText = "No";
            this.clmNO.Name = "clmNO";
            this.clmNO.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.clmNO.Width = 56;
            // 
            // clmDealName
            // 
            this.clmDealName.HeaderText = "Offerte (Deal)";
            this.clmDealName.Name = "clmDealName";
            this.clmDealName.Width = 129;
            // 
            // frmCategory
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.CancelButton = this.cmdCancel;
            this.ClientSize = new System.Drawing.Size(699, 589);
            this.Controls.Add(this.MaindataGridView);
            this.Controls.Add(this.pnlFooter);
            this.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "frmCategory";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Offerte (Deals)";
            this.Load += new System.EventHandler(this.frmCategory_Load);
            this.pnlFooter.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.MaindataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlFooter;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.DataGridView MaindataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmDealName;
    }
}