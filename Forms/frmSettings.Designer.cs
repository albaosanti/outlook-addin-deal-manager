
namespace SajjuCode.OutlookAddIns
{
    partial class frmSettings
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
            this.gbContent = new System.Windows.Forms.GroupBox();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdMessageSaveIn = new System.Windows.Forms.Button();
            this.MsgSaveIntextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmdBrowseDeals = new System.Windows.Forms.Button();
            this.OfferCSVtextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.gbContent.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbContent
            // 
            this.gbContent.Controls.Add(this.cmdCancel);
            this.gbContent.Controls.Add(this.cmdSave);
            this.gbContent.Controls.Add(this.cmdMessageSaveIn);
            this.gbContent.Controls.Add(this.MsgSaveIntextBox);
            this.gbContent.Controls.Add(this.label2);
            this.gbContent.Controls.Add(this.cmdBrowseDeals);
            this.gbContent.Controls.Add(this.OfferCSVtextBox);
            this.gbContent.Controls.Add(this.label1);
            this.gbContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gbContent.Location = new System.Drawing.Point(0, 0);
            this.gbContent.Name = "gbContent";
            this.gbContent.Size = new System.Drawing.Size(579, 167);
            this.gbContent.TabIndex = 0;
            this.gbContent.TabStop = false;
            // 
            // cmdCancel
            // 
            this.cmdCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Location = new System.Drawing.Point(333, 114);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(117, 41);
            this.cmdCancel.TabIndex = 7;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            // 
            // cmdSave
            // 
            this.cmdSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdSave.Location = new System.Drawing.Point(456, 114);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(117, 41);
            this.cmdSave.TabIndex = 6;
            this.cmdSave.Text = "Save";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdMessageSaveIn
            // 
            this.cmdMessageSaveIn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdMessageSaveIn.BackgroundImage = global::SajjuCode.OutlookAddIns.Properties.Resources.folder;
            this.cmdMessageSaveIn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.cmdMessageSaveIn.FlatAppearance.BorderSize = 0;
            this.cmdMessageSaveIn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdMessageSaveIn.Location = new System.Drawing.Point(544, 60);
            this.cmdMessageSaveIn.Name = "cmdMessageSaveIn";
            this.cmdMessageSaveIn.Size = new System.Drawing.Size(29, 29);
            this.cmdMessageSaveIn.TabIndex = 5;
            this.cmdMessageSaveIn.UseVisualStyleBackColor = true;
            this.cmdMessageSaveIn.Click += new System.EventHandler(this.cmdMessageSaveIn_Click);
            // 
            // MsgSaveIntextBox
            // 
            this.MsgSaveIntextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MsgSaveIntextBox.Location = new System.Drawing.Point(116, 59);
            this.MsgSaveIntextBox.Name = "MsgSaveIntextBox";
            this.MsgSaveIntextBox.Size = new System.Drawing.Size(422, 29);
            this.MsgSaveIntextBox.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label2.Location = new System.Drawing.Point(6, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 29);
            this.label2.TabIndex = 3;
            this.label2.Text = "Msg Save In:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmdBrowseDeals
            // 
            this.cmdBrowseDeals.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdBrowseDeals.BackgroundImage = global::SajjuCode.OutlookAddIns.Properties.Resources.csv;
            this.cmdBrowseDeals.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.cmdBrowseDeals.FlatAppearance.BorderSize = 0;
            this.cmdBrowseDeals.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdBrowseDeals.Location = new System.Drawing.Point(544, 25);
            this.cmdBrowseDeals.Name = "cmdBrowseDeals";
            this.cmdBrowseDeals.Size = new System.Drawing.Size(29, 29);
            this.cmdBrowseDeals.TabIndex = 2;
            this.cmdBrowseDeals.UseVisualStyleBackColor = true;
            this.cmdBrowseDeals.Click += new System.EventHandler(this.cmdBrowseDeals_Click);
            // 
            // OfferCSVtextBox
            // 
            this.OfferCSVtextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OfferCSVtextBox.Location = new System.Drawing.Point(116, 25);
            this.OfferCSVtextBox.Name = "OfferCSVtextBox";
            this.OfferCSVtextBox.Size = new System.Drawing.Size(422, 29);
            this.OfferCSVtextBox.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label1.Location = new System.Drawing.Point(6, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "Offerte CSV:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(579, 167);
            this.ControlBox = false;
            this.Controls.Add(this.gbContent);
            this.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "frmSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Configurazione delle offerte";
            this.Load += new System.EventHandler(this.frmSettings_Load);
            this.gbContent.ResumeLayout(false);
            this.gbContent.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbContent;
        private System.Windows.Forms.TextBox OfferCSVtextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cmdBrowseDeals;
        private System.Windows.Forms.TextBox MsgSaveIntextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button cmdMessageSaveIn;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}