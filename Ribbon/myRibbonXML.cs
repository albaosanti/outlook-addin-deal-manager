using SajjuCode.OutlookAddIns.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new myRibbonXML();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace SajjuCode.OutlookAddIns
{
    [ComVisible(true)]
    public class myRibbonXML : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public myRibbonXML()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            //if (ribbonID == "Microsoft.Outlook.Mail.Compose")
            //{
            //    return "";
            //}
            if (ribbonID.ToLower() == "Microsoft.Outlook.Explorer".ToLower())
            {
                return GetResourceText("SajjuCode.OutlookAddIns.Ribbon.myRibbonXML.xml");
            }
            else
            {
                return "";
            }


        }

        #endregion

        #region "Context"

        //public string RemoveSpecialCharacters(string myInput)
        //{
        //    try
        //    {
        //        string myOutPut = myInput;

        //        if (string.IsNullOrEmpty(myOutPut)) return "";
        //        // if (!string.IsNullOrEmpty(mySelectedDeals) && mySelectedDeals.ToLower().Contains(DealText.ToLower().Replace(",", "_").Replace("\""," ").Trim()))

        //        myOutPut = myOutPut.Replace("&", " ");
        //        myOutPut = myOutPut.Replace(",", "_");
        //        myOutPut = myOutPut.Replace("\"", " ");
        //        //myOutPut = myOutPut.Replace(";", " ");

        //        return myOutPut;
        //    }
        //    catch (Exception ex)
        //    {

        //        return myInput;
        //    }
        //}
        
        public string GetMenuContent(Office.IRibbonControl control)
        {
            bool isTask = false;
            bool isDealExist = false;
            StringBuilder xmlString = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" >");
            string mySelectedDeals;
            
            if (Globals.ThisAddIn.ADDIN_INITIALIZE_SUCCESS)
            {
              mySelectedDeals = this.getEmail(ref control, out isTask);
              mySelectedDeals = Globals.ThisAddIn.Deals.RemoveSpecialCharacters(mySelectedDeals);

              Globals.ThisAddIn.Deals.LoadDeals();

              if (Globals.ThisAddIn.Deals.GetDealCount() > 0)
              {
                if (!string.IsNullOrEmpty(mySelectedDeals))
                {
                  foreach (var myD in mySelectedDeals.Split(','))
                  {
                    if (Globals.ThisAddIn.Deals.DealExist(myD))
                    {
                      isDealExist = true;
                      break;
                    }
                  }

                  if (!isDealExist && isTask) return "";
                }
                else
                  if (isTask) return "";

                try
                {
                  string checked_visible = "";
                  foreach (Deal deal in Globals.ThisAddIn.Deals.GetVisibleDeals())
                  {
                    checked_visible = "";
                    //if (!string.IsNullOrEmpty(mySelectedDeals) && mySelectedDeals.ToLower().Contains(DealText.ToLower().Replace(",", "_").Replace("\"", " ").Trim()))
                    if (!string.IsNullOrEmpty(mySelectedDeals) && mySelectedDeals.Replace("_ Offerte", "").ToLower().Trim() == (deal.Name.ToLower().Replace(",", "_").Replace("\"", " ").Trim()).ToLower())
                      checked_visible = "imageMso='ActiveXCheckBox' showImage='true'";

                    xmlString.Append($"<button id=\"nwtDeal_{deal.Index.ToString()}\" label=\"{deal.Name.Replace("\"", " ")}\" tag=\"{ deal.Name.Replace("\"", " ")}\" {checked_visible} onAction=\"OnMyButtonClick\"/>");
                  }
                }
                catch (Exception ex) { }
              }
            }
            
            xmlString.Append(@"</menu>");
            return xmlString.ToString();
        }
       
        #endregion

        #region "Export"

        public string getDealNumber(string CategoryName)
        {
            string FileData = "";
            bool isDealExist = false;
            string FilePath = SettingsClassFunctions.OfferCSVPath;// "C:\\Deals\\deal list.csv";            
            string returnNumber = "00";

            if (File.Exists(FilePath))
            {
                FileData = File.ReadAllText(FilePath);

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

                        if (string.IsNullOrEmpty(myDealLine))
                        {
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

                        if (!string.IsNullOrEmpty(CategoryName) && CategoryName.ToLower().Contains(DealText.ToLower().Replace(",", "_")))
                        {
                            return DealNo;
                        }



                    }
                }
                catch (Exception ex)
                {


                }

            }
            return returnNumber;
            //return "";
        }

        public void ExportMessage(Microsoft.Office.Interop.Outlook.MailItem mailItem, string CategoryName, bool IgnoreUnknown = false)
        {
            try
            {
                if (string.IsNullOrEmpty(CategoryName)) return;

                if (mailItem != null && !string.IsNullOrEmpty(CategoryName))
                {
                    Microsoft.Office.Interop.Outlook.Conversation conversation;
                    Microsoft.Office.Interop.Outlook.Store myStore;
                    conversation = mailItem.GetConversation();
                    if (conversation != null)
                    {
                        myStore = Globals.ThisAddIn.Application.Session.DefaultStore;
                        if (Directory.Exists(SettingsClassFunctions.MessageSaveFolder))
                        {
                            string myNumber = this.getDealNumber(CategoryName.Replace(",", "_"));

                            if (myNumber == "00" || string.IsNullOrEmpty(myNumber))
                            {
                                if (IgnoreUnknown == false)
                                {
                                    mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + this.getDealNumber(CategoryName.Replace(",", "_")) + " - " + conversation.ConversationID + ".msg",
                                Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                }
                            }
                            else
                            {
                                mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + this.getDealNumber(CategoryName.Replace(",", "_")) + " - " + conversation.ConversationID + ".msg",
                                Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                            }


                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }


        }

        public void cmdExportDealClick(Office.IRibbonControl control)
        {

            var myC = Globals.ThisAddIn.Application.ActiveExplorer().Selection;

            if (myC.Count > 0)
            {
                var selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selObject is Microsoft.Office.Interop.Outlook.MailItem)
                {

                    Microsoft.Office.Interop.Outlook.MailItem mailItem = (selObject as Microsoft.Office.Interop.Outlook.MailItem);
                    if (mailItem != null)
                    {
                        //MainClassFunctions.ExportAllConversationMessage(mailItem);

                        //if (!string.IsNullOrEmpty(mailItem.Categories))
                        //{
                        //    foreach (string myCat in mailItem.Categories.Split(','))
                        //    {
                        //        if (myCat.ToLower().Trim() == "Offerte".ToLower())
                        //        {
                        //            continue;
                        //        }
                        //        this.ExportMessage(mailItem, myCat, true);
                        //    }
                        //}

                    }
                }

            }




        }
        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

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

        public string getEmail(ref Office.IRibbonControl control,out Boolean isTask)
        {
            isTask = false;
            try
            {
                if(control.Context is Outlook.Selection)
                {
                   Outlook.Selection selection = (Outlook.Selection)control.Context;
                   if(selection.Count > 0)
                   {
                      if (selection[1] is Outlook.MailItem)
                      {
                        return ((Outlook.MailItem)selection[1]).Categories;
                      }
                      else if ( selection[1] is Outlook.TaskItem)
                      {
                        isTask = true;
                        return ((Outlook.TaskItem)selection[1]).Categories;
                      }
                   }
                }
                return "";
            }
            catch (Exception ex)
            {
                return "";
            }

        }

        public void OnMyButtonClick(Office.IRibbonControl control)
        {
          if (control.Context is Outlook.Selection)
          {
            var selected = (Outlook.Selection)control.Context;
            if (control.Tag !=null && selected.Count > 0 )
            {
              if (selected[1] is Outlook.MailItem)
              {
                Outlook.MailItem mailitem = (Outlook.MailItem)selected[1];
                Item_AssignCategory(mailitem, control.Tag);
              }
              else if(selected[1] is Outlook.TaskItem)
              {
                Outlook.TaskItem taskitem = (Outlook.TaskItem)selected[1];
                Task_AssignCategory(taskitem, control.Tag);
              }
            }
          }
        }

        public void Item_AssignCategory(Microsoft.Office.Interop.Outlook.MailItem mailItem, string CategoryName)
        {
            try
            {
                if (mailItem != null && !string.IsNullOrEmpty(CategoryName))
                {
                    Outlook.Conversation conversation;
                    Outlook.Store myStore;
                    conversation = mailItem.GetConversation();
                    if (conversation != null)
                    {
                        myStore = Globals.ThisAddIn.Application.Session.DefaultStore;

                        conversation.ClearAlwaysAssignCategories(myStore);
                        conversation.SetAlwaysAssignCategories("Offerte", myStore);
                        conversation.SetAlwaysAssignCategories(CategoryName.Replace(",", "_"), myStore);

                        Globals.ThisAddIn.AddConversation_To_Export(conversation, CategoryName);                        
                    }
                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
      }

      public void Task_AssignCategory(Microsoft.Office.Interop.Outlook.TaskItem mailItem, string CategoryName)
        {
            try
            {
                if (mailItem != null && !string.IsNullOrEmpty(CategoryName))
                {
                    //Microsoft.Office.Interop.Outlook.Conversation conversation;
                    //Microsoft.Office.Interop.Outlook.Store myStore;
                    //conversation = mailItem.GetConversation();
                    //if (conversation != null)
                    //{
                    //    myStore = Globals.ThisAddIn.Application.Session.DefaultStore;
                    //    conversation.ClearAlwaysAssignCategories(myStore);
                    //    conversation.SetAlwaysAssignCategories(CategoryName, myStore);

                    //}
                    mailItem.Categories = "Offerte," + CategoryName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void OncmdContRemoveDeals(Office.IRibbonControl control)
        {
            try
            {
                if (MessageBox.Show("Vuoi concludere le trattative?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    return;
                }

                var myC = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
                if (myC.Count > 0)
                {
                    var selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                    if (selObject is Microsoft.Office.Interop.Outlook.MailItem)
                    {
                        Microsoft.Office.Interop.Outlook.MailItem mailItem = (selObject as Microsoft.Office.Interop.Outlook.MailItem);
                        if (mailItem != null)
                        {
                            Microsoft.Office.Interop.Outlook.Conversation conversation;
                            Microsoft.Office.Interop.Outlook.Store myStore;
                            conversation = mailItem.GetConversation();
                            if (conversation != null)
                            {
                                myStore = Globals.ThisAddIn.Application.Session.DefaultStore;
                                conversation.ClearAlwaysAssignCategories(myStore);
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {


            }

        }

        public void OnShowCategoryClick(Office.IRibbonControl control)
        {

            try
            {
                if(control.Context is Outlook.Selection)
                {
                  Outlook.Selection selected = (Outlook.Selection) control.Context;
                  if(selected.Count > 0)
                  {
                    if(selected[1] is Outlook.MailItem)
                    {
                       Outlook.MailItem mailItem = (Outlook.MailItem)selected[1];
                       mailItem.ShowCategoriesDialog();
                    }
                  }
                }

                //var myC = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
                //if (myC.Count > 0)
                //{
                //    var selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                //    if (selObject is Microsoft.Office.Interop.Outlook.MailItem)
                //    {
                //        Microsoft.Office.Interop.Outlook.MailItem mailItem = (selObject as Microsoft.Office.Interop.Outlook.MailItem);
                //        if (mailItem != null)
                //        {
                //            mailItem.ShowCategoriesDialog();
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {


            }
        }       

        public void OncmdCreateDealsClick(Office.IRibbonControl control)
        {
            try
            {
              if (Globals.ThisAddIn.ADDIN_INITIALIZE_SUCCESS)
              {
                if (control.Context is Outlook.Selection)
                {
                  Outlook.Selection selected = (Outlook.Selection)control.Context;
                  if (selected.Count > 0)
                  {
                    if (selected[1] is Outlook.MailItem)
                    {
                      string DealName = "";
                      string FullDomainName = "";
                      string DomainName = "";
                      int i = 0;
                      Outlook.MailItem mailItem = (Outlook.MailItem)selected[1];

                      foreach (string myS in mailItem.SenderEmailAddress.Split('@'))
                      {
                        if (i == 1)
                        {
                          FullDomainName = myS;
                          DomainName = FullDomainName.Substring(0, myS.LastIndexOf("."));
                        }
                        i++;
                      }

                      DealName = (mailItem.SenderName.Replace(FullDomainName, "").Replace("@", "") + " " + DomainName.ToUpper() + " " + mailItem.Subject);//.ToUpper();
                                                                                                                                                          //SaveDeal(DealName);
                      Globals.ThisAddIn.Deals.AppendNewDeal(DealName);
                      Item_AssignCategory(mailItem, DealName);
                    }
                  }
                }
              }
            }catch (Exception ex)
            {}
        }

        public void OncmdCreateTaskDealsClick(Office.IRibbonControl control)
        {
          try
          {
            if (control.Context is Outlook.Selection)
            {
              Outlook.Selection selected = (Outlook.Selection)control.Context;
              if (selected[1] is Outlook.MailItem)
              {
                if (selected.Count > 0)
                {
                  if (selected[1] is Outlook.TaskItem)
                  {
                    Outlook.TaskItem taskItem = (Outlook.TaskItem)selected[1];
                    int i = 0;
                    string DealName = "";
                    string FullDomainName = "";
                    string DomainName = "";

                    foreach (string myS in taskItem.Owner.Split('@'))
                    {
                      if (i == 1)
                      {
                        FullDomainName = myS;
                        DomainName = FullDomainName.Substring(0, myS.LastIndexOf("."));
                      }
                      i++;
                    }

                    DealName = (taskItem.Owner.Replace(FullDomainName, "").Replace("@", "") + " " + DomainName.ToUpper() + " " + taskItem.Subject);//.ToUpper();
                    //SaveDeal(DealName);
                    Globals.ThisAddIn.Deals.AppendNewDeal(DealName);
                    Task_AssignCategory(taskItem, DealName);
                  }
                }
              }
            }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message);}
        }

        void SaveDeal(string DealName)
        {
            try
            {
                frmCategory frmCategory = new frmCategory();
                frmCategory.validateFileFolder();
                string LastLine = "";
                int LastNumber = 1;
                foreach (string myLine in File.ReadAllLines(frmCategory.FilePath))
                {
                    LastLine = myLine;

                    if (myLine.ToLower().Contains(DealName.ToLower()+";"))
                    {
                        MessageBox.Show("Esiste già un affare con lo stesso nome!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                if (!string.IsNullOrEmpty(LastLine))
                {
                    foreach (string myS in LastLine.Split(';'))
                    {
                        int.TryParse(myS, out LastNumber);
                        break;
                    }
                }

                string myData = File.ReadAllText(frmCategory.FilePath);
                try
                {
                    int i = 0;
                    foreach (var myS in myData.Split(new[] { '\r', '\n' }))
                    {
                        if (i > 0)
                        {
                            break;
                        }

                        myData = myData.Replace(myS, "Index;Deal name;Visible;Section");

                        i++;
                    }
                }
                catch (Exception ex)
                { MessageBox.Show(ex.Message); }
        myData = myData + Environment.NewLine + (LastNumber + 1) + ";" + DealName + ";1;1";

                File.WriteAllText(frmCategory.FilePath, myData);
            }
            catch (Exception ex)
            {


            }
        }

        public void OncmdDealsClick(Office.IRibbonControl control)
        {

            try
            {
                new frmCategory().ShowDialog();
            }
            catch (Exception ex)
            {
            }
        }

        public void OncmdSettingsClick(Office.IRibbonControl control)
        {
            try
            {
                if(new frmSettings().ShowDialog() == DialogResult.OK)
                {
                  if (!Globals.ThisAddIn.ADDIN_INITIALIZE_SUCCESS)
                  {
                    Globals.ThisAddIn.ADDIN_INITIALIZE_SUCCESS = true;
                    Globals.ThisAddIn.InitGlobalVariables();
                  }
                  else
                  {
                    Globals.ThisAddIn.Deals.SourceCSV = SettingsClassFunctions.OfferCSVPath;
                    Globals.ThisAddIn.Deals.LoadDeals(true);
                    Globals.ThisAddIn.Logs.LogPath = SettingsClassFunctions.OfferCSVPath + "\\Log\\";
                  }
                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }


        private void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.Selection Selection)
        {

        }
        #endregion
    }
}
