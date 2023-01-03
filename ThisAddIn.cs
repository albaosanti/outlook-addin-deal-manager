using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SajjuCode.OutlookAddIns.Base;
using System.Reflection;
using System.IO;
using System.ComponentModel;

namespace SajjuCode.OutlookAddIns
{
    public partial class ThisAddIn
    {
        public bool ADDIN_INITIALIZE_SUCCESS = false;
        public int NXT_EXPORT_INTERVAL_SECONDS = 3;
        public BackgroundWorker bgwExport;
        public Timer myTimer;
        public DateTime NextRun;

        public NameSpace outlookNameSpace;
        public MAPIFolder inbox;
        public MAPIFolder sentfolder;
        public Items inbox_items;
        public Items sent_items;
    
        public DealManager Deals = new DealManager();
        public List<OfferteConversation> ExportQueue = new List<OfferteConversation>();
        public ExportLog Logs = new ExportLog();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)        
        {
          //INITIALIZE SETTINGS
          this.ADDIN_INITIALIZE_SUCCESS = true;
          if (SettingsClassFunctions.isDealFileExists == false  || !SettingsClassFunctions.ReadSettingFile())
          {
            MessageBox.Show("Setting file not found or some settings are missing!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (new frmSettings().ShowDialog() != DialogResult.OK)
            {
              MessageBox.Show("Ci sono configurazioni mancanti!", "Outlook AddIns - Offerta");
              this.ADDIN_INITIALIZE_SUCCESS = false;
            }
          }
          //INITIALIZE VARIABLES
          if (this.ADDIN_INITIALIZE_SUCCESS)
          {           
            InitGlobalVariables();           
          }
        }

        #region "Background Exporting Process"
        public void InitGlobalVariables()
        {
            MainClassFunctions.Application = this.Application;
            MainClassFunctions.objExplorer = MainClassFunctions.Application.Explorers;
            MainClassFunctions.AddACategory("Offerte");
            //AddACategory();
            //RemoveCategory("");

          #region "Item_Added Event - Inbox & Sent Folder"
            outlookNameSpace = this.Application.GetNamespace("MAPI");

            inbox = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            sentfolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);

            inbox_items = inbox.Items;
            sent_items = sentfolder.Items;

            inbox_items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
            sent_items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
          #endregion

          this.Logs = new ExportLog();
          this.Logs.LogPath = Path.GetDirectoryName(SettingsClassFunctions.OfferCSVPath) + "\\Logs\\";

          this.Deals.SourceCSV = SettingsClassFunctions.OfferCSVPath;
          this.Deals.LoadDeals(true);

          this.ExportQueue = new List<OfferteConversation>();

          this.myTimer = new Timer();
          this.myTimer.Interval = 1000;//3 SECONDS
          this.myTimer.Enabled = false;
          this.myTimer.Tick += MyTimer_Tick;

          this.bgwExport = new BackgroundWorker();
          this.bgwExport.WorkerSupportsCancellation = true;
          this.bgwExport.WorkerReportsProgress = true;
          this.bgwExport.DoWork += Worker_DoWork;
          this.bgwExport.ProgressChanged += Worker_ProgressChanged;
          this.bgwExport.RunWorkerCompleted += Worker_RunWorkerCompleted;

          if (!Directory.Exists(SettingsClassFunctions.MessageSaveFolder))
            Directory.CreateDirectory(SettingsClassFunctions.MessageSaveFolder);
          
          //run the exporting timer & backgroundworkers
          NextRun = DateTime.Now;
          myTimer.Enabled = true;
          myTimer.Start();
        }
      
        public bool ExportQueue_Add(Base.OfferteConversation conversation)
        {
          bool added = false;
          if (this.ADDIN_INITIALIZE_SUCCESS)
          {
            if (ExportQueue.FindAll(co => co.ConversationID.ToLower() == conversation.ConversationID.ToLower()).Count == 0)
            {
              ExportQueue.Add(conversation);
              added = true;
            }

          }
          return added;
        }

        private void Worker_DoWork(object sender,DoWorkEventArgs e)
          {
          OfferteConversation offerte_convo;
          string targetFilename;
        
          if (!this.ADDIN_INITIALIZE_SUCCESS) return;
          
          try
          {
            while (ExportQueue.Count > 0)
            {
              //FIRST IN FIRST OUT EXPORT
              offerte_convo = ExportQueue[0];
              ExportQueue.RemoveAt(0);

              //PROCESS CONVERSATION HERE
              Outlook.SimpleItems simpleItems = offerte_convo.conversation.GetRootItems();
              if (simpleItems != null)
              {
                foreach (object conv_itm in simpleItems)
                {
                  if (conv_itm is Outlook.MailItem)
                  {
                    targetFilename = "";
                    try
                    {
                      Outlook.MailItem i_tm = (Outlook.MailItem)conv_itm;
                      targetFilename = i_tm.SentOn.ToString("yyyyMMddHHmmss");
                      targetFilename += $" {Deals.RemoveSpecialCharacters(i_tm.SenderName).Trim()}";
                      targetFilename += $" ({offerte_convo.DealIndex})";

                      if (!File.Exists($"{SettingsClassFunctions.MessageSaveFolder}\\{targetFilename}.msg"))
                      {
                        i_tm.SaveAs($"{SettingsClassFunctions.MessageSaveFolder}\\{targetFilename}.msg");
                        Logs.WriteLog($"Success. {targetFilename}.msg",false);
                      }
                    }
                    catch (System.Exception e1) { Logs.WriteLog($"Error: {targetFilename} {e1.Message}", true);}                
                  }
                  EnumerateConversation(conv_itm, offerte_convo.conversation,offerte_convo.DealIndex);
                }
              }
            }
          }
          catch (System.Exception ex) { Logs.WriteLog($"Error@Worker: {ex.Message}", true); }
        }

        public void EnumerateConversation(object item,Outlook.Conversation conversation,string deal_index)
        {
          Outlook.SimpleItems items = conversation.GetChildren(item);
          if (items.Count > 0)
          {
            string targetFilename;
            foreach (object myItem in items)
            {
              // In this example, enumerate only MailItem type.
              // Other types such as PostItem or MeetingItem
              // can appear in Conversation.
              if (myItem is Outlook.MailItem)
              {
                targetFilename = "";
                try
                {
                  Outlook.MailItem i_tm = (Outlook.MailItem)myItem;
                  targetFilename = i_tm.SentOn.ToString("yyyyMMddHHmmss");
                  targetFilename += $" {Deals.RemoveSpecialCharacters(i_tm.SenderName).Trim()}";
                  targetFilename += $" ({deal_index})";

              if (!File.Exists($"{SettingsClassFunctions.MessageSaveFolder}\\{targetFilename}.msg"))
                  {
                    i_tm.SaveAs($"{SettingsClassFunctions.MessageSaveFolder}\\{targetFilename}.msg");
                    Logs.WriteLog($"Success. {targetFilename}.msg", false);
                  }
                }
                catch (System.Exception e) { Logs.WriteLog($"Error: {targetFilename}.msg {e.Message}", true); }
              }
              // Continue recursion.
              EnumerateConversation(myItem, conversation, deal_index);
            }
          }
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {}

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
          NextRun = DateTime.Now.AddSeconds(NXT_EXPORT_INTERVAL_SECONDS);
        }

        private void MyTimer_Tick(object sender, EventArgs e)
        {
          myTimer.Enabled = false;
          if (!this.ADDIN_INITIALIZE_SUCCESS)
            return;
          else
          {
            if (DateTime.Now >= NextRun)
            {
              if (!bgwExport.IsBusy && ExportQueue.Count > 0)
                bgwExport.RunWorkerAsync();
            }
          }
          myTimer.Enabled = true;
        }

        public void AddConversation_To_Export(Outlook.Conversation conversation, string CategoryName)
      {
        if (!this.ADDIN_INITIALIZE_SUCCESS) return;
        OfferteConversation offerte_conversation = new OfferteConversation(conversation.ConversationID, conversation);
        List<Deal> match_deal = Globals.ThisAddIn.Deals.GetMatchDeal(CategoryName);
        Deal deal = null;

        if (match_deal.Count == 0)
          deal = Globals.ThisAddIn.Deals.AppendNewDeal(CategoryName);
        else
          deal = match_deal[0];

        if (deal != null)
          offerte_conversation.DealIndex = deal.Index.ToString();

        //ADD TO EXPORT QUEUE
        Globals.ThisAddIn.ExportQueue_Add(offerte_conversation);
      }
    #endregion

        #region "Handlers"
    private static string GetResourceText(string resourceName)
    {
      Assembly asm = Assembly.GetExecutingAssembly();
      string[] resourceNames = asm.GetManifestResourceNames();
      for (int i = 0; i < resourceNames.Length; ++i)
      {
        if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
        {
          using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
          {
            if (resourceReader != null)
            {
              return resourceReader.ReadToEnd();
            }
          }
        }
      }
      return null;
    }

    void Items_ItemAdd(object item)
    {
      try
      {
        if (item != null && this.ADDIN_INITIALIZE_SUCCESS)
        {
          if (item is Outlook.MailItem)
          {
            Outlook.MailItem mail = (Outlook.MailItem)item;
            if (mail.Categories != null)
            {
              string categories = mail.Categories.Replace(", ", "");
              if (categories.ToLower().Contains("offerte"))
              {
                //find the deal 
                foreach (Deal deal in Deals.GetVisibleDeals())
                {
                  if (Deals.RemoveSpecialCharacters(categories.ToLower()).Contains(Deals.RemoveSpecialCharacters(deal.Name.ToLower())))
                  {
                    AddConversation_To_Export(mail.GetConversation(), deal.Name);
                    break;
                  }
                }
              }
            }
          }
        }
      }
      catch (System.Exception ex) { Logs.WriteLog($"@Items_ItemAdd: {ex.Message}", true); }
    }

    private void Application_ViewContextMenuDisplay(Office.CommandBar CommandBar, Outlook.View View)
    {
      "a".ToString();
    }

    private void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
    {
      if (Selection[1] is MailItem)
      {

        MailItem selectedMailItem = Selection[1] as MailItem;

        this.CustomContextMenu(CommandBar, Selection);

      }
    }

    private void CustomContextMenu(Office.CommandBar CommandBar, Outlook.Selection Selection)
    {

      Office.CommandBarButton customContextMenuTag = (Office.CommandBarButton)CommandBar.Controls.Add
      (Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);

      customContextMenuTag.Click += new

      Office._CommandBarButtonEvents_ClickEventHandler(customContextMenuTag_Click);

      customContextMenuTag.Caption = "Assign Deal";

      //customContextMenuTag.FaceId = 351; //displays the image for the menu item

      customContextMenuTag.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;

    }

    private void customContextMenuTag_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
    {

    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
      // Note: Outlook no longer raises this event. If you have code that 
      //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
    }
    #endregion



    #region VSTO generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);;
        }

        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new myRibbonXML();
        }


        #region "Category"
        private void EnumerateCategories()
        {
            Outlook.Categories categories = Application.Session.Categories;
            foreach (Outlook.Category category in categories)
            {
                //Debug.WriteLine(category.Name);
                //Debug.WriteLine(category.CategoryID);
            }
        }
        #endregion
    }
}
