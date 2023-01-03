
namespace SajjuCode.OutlookAddIns
{
    partial class myRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public myRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.Maintab = this.Factory.CreateRibbonTab();
			this.Dealgroup = this.Factory.CreateRibbonGroup();
			this.NewDealbutton = this.Factory.CreateRibbonButton();
			this.CloseDealbutton = this.Factory.CreateRibbonButton();
			this.separator1 = this.Factory.CreateRibbonSeparator();
			this.Maintab.SuspendLayout();
			this.Dealgroup.SuspendLayout();
			this.SuspendLayout();
			// 
			// Maintab
			// 
			this.Maintab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.Maintab.Groups.Add(this.Dealgroup);
			this.Maintab.Label = "TabAddIns";
			this.Maintab.Name = "Maintab";
			// 
			// Dealgroup
			// 
			this.Dealgroup.Items.Add(this.NewDealbutton);
			this.Dealgroup.Items.Add(this.CloseDealbutton);
			this.Dealgroup.Items.Add(this.separator1);
			this.Dealgroup.Label = "Deals";
			this.Dealgroup.Name = "Dealgroup";
			// 
			// NewDealbutton
			// 
			this.NewDealbutton.Image = global::SajjuCode.OutlookAddIns.Properties.Resources.check;
			this.NewDealbutton.Label = "New Deal";
			this.NewDealbutton.Name = "NewDealbutton";
			this.NewDealbutton.ShowImage = true;
			// 
			// CloseDealbutton
			// 
			this.CloseDealbutton.Label = "Close Deal";
			this.CloseDealbutton.Name = "CloseDealbutton";
			// 
			// separator1
			// 
			this.separator1.Name = "separator1";
			// 
			// myRibbon
			// 
			this.Name = "myRibbon";
			this.RibbonType = "Microsoft.Outlook.Mail.Read";
			this.Tabs.Add(this.Maintab);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.myRibbon_Load);
			this.Maintab.ResumeLayout(false);
			this.Maintab.PerformLayout();
			this.Dealgroup.ResumeLayout(false);
			this.Dealgroup.PerformLayout();
			this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Maintab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Dealgroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewDealbutton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CloseDealbutton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal myRibbon myRibbon
        {
            get { return this.GetRibbon<myRibbon>(); }
        }
    }
}
