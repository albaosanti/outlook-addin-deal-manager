using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Exception = System.Exception;

namespace SajjuCode.OutlookAddIns.Base
{
    public class MainClassFunctions
    {
        public static Explorers objExplorer { get; set; }
        public static MailItem objMail { get; set; }
        public static Microsoft.Office.Interop.Outlook.Application Application { get; set; }

        public static void setEvents()
        {
            try
            {
                objExplorer.NewExplorer += ObjExplorer_NewExplorer;

            }
            catch (System.Exception ex)
      { MessageBox.Show(ex.Message); }
    }

        private static void ObjExplorer_NewExplorer(Explorer Explorer)
        {
            try
            {
                Explorer.ToString();
            }
            catch (System.Exception ex)
      { MessageBox.Show(ex.Message); }
    }

        #region "Category"
        public static void AddACategory(string categoryLabel)
        {
            Microsoft.Office.Interop.Outlook.Categories categories = Application.Session.Categories;
            if (!CategoryExists(categoryLabel))
            {
                Microsoft.Office.Interop.Outlook.Category category = categories.Add(categoryLabel, Microsoft.Office.Interop.Outlook.OlCategoryColor.olCategoryColorDarkBlue);
            }

            //CreateTextAndCategoryRule();
            //AssignCategories();
        }

        public static bool CategoryExists(string categoryName)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Category category = Application.Session.Categories[categoryName];
                if (category != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch { return false; }
        }


        public static void CreateTextAndCategoryRule(string categoryName, string SubjectContain)
        {
            //if (!CategoryExists(categoryName))
            //{
            //    Application.Session.Categories.Add(categoryName, Type.Missing, Type.Missing);
            //}

            //if (!CategoryExists("Deals"))
            //{
            //    Application.Session.Categories.Add("Deals", Type.Missing, Type.Missing);
            //}

            Boolean isRuleExist = false;
            Microsoft.Office.Interop.Outlook.Rules rules = Application.Session.DefaultStore.GetRules();
            Microsoft.Office.Interop.Outlook.Rule textRule = rules.Create("Deal Rule " + categoryName, Microsoft.Office.Interop.Outlook.OlRuleType.olRuleReceive);
            if (rules != null && rules.Count > 0)
            {
                try
                {
                    for (int i = 1; i <= rules.Count; i++)
                    {
                        Microsoft.Office.Interop.Outlook.Rule rule = rules[i];
                        if (rule != null && rule.Name.ToLower() == ("Deal Rule " + categoryName).ToLower())
                        {
                            textRule = rule;
                        }
                    }

                }
                catch (System.Exception ex)
        { MessageBox.Show(ex.Message); }
      }

            if (textRule == null)
            {
                return;
            }


            Object[] textCondition = { SubjectContain };
            Object[] categoryAction = { "Deals", categoryName };
            textRule.Conditions.BodyOrSubject.Text = textCondition;
            textRule.Conditions.BodyOrSubject.Enabled = true;
            textRule.Actions.AssignToCategory.Categories = categoryAction;
            textRule.Actions.AssignToCategory.Enabled = true;
            rules.Save(true);

        }

        public static void AssignCategories(Microsoft.Office.Interop.Outlook.MailItem mailItem, string categoryName, int itemIndex = 0)
        {
            //string filter = "@SQL=" + "\"" + "urn:schemas:httpmail:subject"
            //    + "\"" + " ci_phrasematch 'ISV'";

            //if (!CategoryExists(categoryName))
            //{
            //    Application.Session.Categories.Add(categoryName, Type.Missing, Type.Missing);
            //}

            //if (!CategoryExists("Deals"))
            //{
            //    Application.Session.Categories.Add("Deals", Type.Missing, Type.Missing);
            //}

            if (mailItem != null)
            {
                mailItem.Categories = categoryName + ",Deals";
                mailItem.Save();
            }
            //var myII = Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox).Items;

            ////Outlook.Items items =Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items.Restrict(filter);
            //for (int i = 1; i <= 5; i++)
            //{
            //    OutlookItem item = new OutlookItem(myII[i]);
            //    string existingCategories = item.Categories;
            //    //if (String.IsNullOrEmpty(existingCategories))
            //    //{
            //    //    item.Categories = categoryName;
            //    //}
            //    //else
            //    //{
            //    //    if (item.Categories.Contains(categoryName) == false)
            //    //    {
            //    //        item.Categories = existingCategories + ", " + categoryName;
            //    //    }
            //    //}
            //    item.Categories = categoryName + ",Deals";
            //    item.Save();
            //}
        }


        #endregion "Category"

        #region "Export Message"

        public static string Export_getDealNumber(string CategoryName)
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
                catch (System.Exception ex)
        { MessageBox.Show(ex.Message); }

      }
            return returnNumber;
            //return "";
        }

        public static void ExportMessage(Microsoft.Office.Interop.Outlook.MailItem mailItem, string CategoryName, bool IgnoreUnknown = false,int childid=0)
        {
            try
            {
                if (string.IsNullOrEmpty(CategoryName)) return;
                string FileName = "";

                if (mailItem != null && !string.IsNullOrEmpty(CategoryName))
                {
                    Microsoft.Office.Interop.Outlook.Conversation conversation;
                    Microsoft.Office.Interop.Outlook.Store myStore;
                    //FileName = mailItem.SenderName + "-" + mailItem.CreationTime.ToString("yyyyMMddHHmmss")
                    //           + " - " + mailItem.ConversationID + " - " + childid + " (" + Export_getDealNumber(CategoryName.Replace(",", "_")) + ")" ;
                    FileName = mailItem.SenderName + "-" + mailItem.CreationTime.ToString("yyyyMMddHHmmss")
                              + " - " + mailItem.ConversationID + " (" + Export_getDealNumber(CategoryName.Replace(",", "_")) + ")";
                    conversation = mailItem.GetConversation();
                    if (conversation != null)
                    {
                        myStore = Globals.ThisAddIn.Application.Session.DefaultStore;
                        if (Directory.Exists(SettingsClassFunctions.MessageSaveFolder))
                        {
                            string myNumber = Export_getDealNumber(CategoryName.Replace(",", "_"));

                            if (myNumber == "00" || string.IsNullOrEmpty(myNumber))
                            {
                                if (IgnoreUnknown == false)
                                {
                                    //if (!File.Exists(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + " - " + conversation.ConversationID + ".msg"))
                                    //{
                                    //    mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + " - " + conversation.ConversationID + ".msg",
                                    //Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                    //}

                                    if (!File.Exists(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + ".msg"))
                                    {
                                        mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + ".msg",
                                    Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                    }
                                }
                            }
                            else
                            {
                                //if (!File.Exists(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + " - " + conversation.ConversationID + ".msg"))
                                //{
                                //    mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + " - " + conversation.ConversationID + ".msg",
                                //Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                //}

                                if (!File.Exists(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + ".msg"))
                                {
                                    mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + ".msg",
                                Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                }
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

        public static void ExportAllConversationMessage(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            try
            {

                string FileName = "";
                string ConversationTopic;
                string ParentCategory = "";

                if (mailItem != null)
                {
                    Microsoft.Office.Interop.Outlook.Conversation conversation;
                    Microsoft.Office.Interop.Outlook.Store myStore;
                    //FileName = mailItem.SenderName + "-" + mailItem.CreationTime.ToString("YYYYMMddHHmmss")
                    //           + " (" + Export_getDealNumber(CategoryName.Replace(",", "_")) + ")";

                   
                    //export selected email
                    //if (!string.IsNullOrEmpty(mailItem.Categories))
                    //{
                    //    foreach (string myCat in mailItem.Categories.Split(','))
                    //    {
                    //        if (myCat.ToLower().Trim() == "Offerte".ToLower())
                    //        {
                    //            continue;
                    //        }
                    //        ExportMessage(mailItem, myCat, true);
                    //    }
                    //}

                    conversation = mailItem.GetConversation();
                    if (conversation != null)
                    {
                        SimpleItems OlSimpleItems = mailItem.GetConversation().GetRootItems();
                        int childid = 0;
                        var myP = OlSimpleItems.Parent;
                        foreach (var ri in OlSimpleItems)
                        {
                            childid++;

                            if (ri is MailItem)
                            {
                                MailItem icr = (MailItem)ri;

                                if (!string.IsNullOrEmpty(icr.Categories))
                                {
                                    
                                    foreach (string myCat in icr.Categories.Split(','))
                                    {
                                        if (myCat.ToLower().Trim() == "Offerte".ToLower())
                                        {
                                            continue;
                                        }
                                        ParentCategory = myCat;
                                        ExportMessage(icr, myCat, true, childid);
                                    }
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(ParentCategory))
                                    {
                                        ExportMessage(icr, ParentCategory, true, childid);
                                    }
                                }
                            }

                            SimpleItems OlChildren = mailItem.GetConversation().GetChildren(ri);
                            foreach (object rci in OlChildren)
                            {
                                childid++;
                                if (rci is MailItem)
                                {
                                    MailItem icr = (MailItem)rci;

                                    if (!string.IsNullOrEmpty(icr.Categories))
                                    {
                                        foreach (string myCat in icr.Categories.Split(','))
                                        {
                                            if (myCat.ToLower().Trim() == "Offerte".ToLower())
                                            {
                                                continue;
                                            }
                                            ParentCategory = myCat;
                                            ExportMessage(icr, myCat, true,childid);
                                        }
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(ParentCategory))
                                        {
                                            ExportMessage(icr, ParentCategory, true, childid);
                                        }
                                    }
                                }
                            }
                        }

                        //myChildren.Count.ToString();

                        //myStore = Globals.ThisAddIn.Application.Session.DefaultStore;
                        //if (Directory.Exists(SettingsClassFunctions.MessageSaveFolder))
                        //{
                        //    string myNumber = Export_getDealNumber(CategoryName.Replace(",", "_"));

                        //    if (myNumber == "00" || string.IsNullOrEmpty(myNumber))
                        //    {
                        //        if (IgnoreUnknown == false)
                        //        {
                        //            mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + " - " + conversation.ConversationID + ".msg",
                        //        Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                        //        }
                        //    }
                        //    else
                        //    {
                        //        mailItem.SaveAs(SettingsClassFunctions.MessageSaveFolder + "\\" + FileName + " - " + conversation.ConversationID + ".msg",
                        //        Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                        //    }


                        //}

                    }

                    ExportByFilterByCategory(mailItem.SenderEmailAddress, mailItem.ConversationID, mailItem.ConversationTopic,ParentCategory);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }


        }

        public static void ExportSingleMessage(Microsoft.Office.Interop.Outlook.MailItem mailItem,string ParentCategory)
        {
            try
            {

                string FileName = "";

                if (mailItem != null)
                {
                    Microsoft.Office.Interop.Outlook.Conversation conversation;
                    Microsoft.Office.Interop.Outlook.Store myStore;
                    //FileName = mailItem.SenderName + "-" + mailItem.CreationTime.ToString("YYYYMMddHHmmss")
                    //           + " (" + Export_getDealNumber(CategoryName.Replace(",", "_")) + ")";


                    //export selected email
                    if (!string.IsNullOrEmpty(mailItem.Categories))
                    {
                        foreach (string myCat in mailItem.Categories.Split(','))
                        {
                            if (myCat.ToLower().Trim() == "Offerte".ToLower())
                            {
                                continue;
                            }
                            ExportMessage(mailItem, myCat, true);
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(ParentCategory))
                        {
                            ExportMessage(mailItem, ParentCategory, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }


        }

        public static void ExportByFilterByCategory(string sendTo,string ConversionID, string Subject, string ParentCategory)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.MAPIFolder inbox = Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                Items items = inbox.Items;
                MailItem mailItem = null;
                object folderItem;
                string subjectName = string.Empty;
                string filter = @"[Subject] = """ + Subject.ToLower().Replace("re", "").Replace(":", "").Trim() + "\"";// "[Subject] > '" + CategoryName.ToLower().Replace("re","").Replace(":","").Trim() +"'";
                //(to:"sajjucode@accsoft.com")
                //string filter = @"(to:\"""+sendTo+"\")";
                //filter= @"[SenderEmailAddress] = """ + sendTo + "\"";
                //filter = @"[ReceivedByName] = """ + sendTo + "\"";
                //filter = @"[ConversationTopic ] = """ + subjectName + "\"";

                //filter = "subject:[New Email]";
                //folderItem = items.Find(filter);
                //folderItem = items.Find(@"[Subject] = """ + CategoryName.ToLower().Replace("re", "").Replace(":", "").Trim() + "\"");

                //while (folderItem != null)
                //{
                //    mailItem = folderItem as MailItem;
                //    if (mailItem != null)
                //    {
                //        //subjectName += "\n" + mailItem.Subject;
                //        ExportSingleMessage(mailItem);
                //    }
                //    folderItem = items.FindNext();
                //}
                //subjectName = " The following e-mail messages were found: " +
                //    subjectName;
                //MessageBox.Show(subjectName);


                Microsoft.Office.Interop.Outlook.MAPIFolder sendFolder = Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail);
                items = sendFolder.Items;
                folderItem = items.Find(filter);
                //folderItem = items.Find(@"[Subject] = """ + CategoryName.ToLower().Replace("re", "").Replace(":", "").Trim() + "\"");

                while (folderItem != null)
                {
                    mailItem = folderItem as MailItem;
                    if (mailItem != null)
                    {
                        //subjectName += "\n" + mailItem.Subject;
                        if (mailItem.ConversationTopic.ToLower().Contains(Subject.ToLower()) || mailItem.Subject.ToLower().Contains(Subject.ToLower()))
                        {                            
                            ExportSingleMessage(mailItem, ParentCategory);
                        }
                    }
                    folderItem = items.FindNext();
                }
            }
            catch (Exception ex)
      { MessageBox.Show(ex.Message); }
    }


        public static void ExportSendedEmail(string sendTo, string ConversionID, string Subject)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.MAPIFolder inbox = Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                Items items = inbox.Items;
                MailItem mailItem = null;
                object folderItem;
                string subjectName = string.Empty;
                string filter = @"[Subject] = """ + Subject.ToLower().Replace("re", "").Replace(":", "").Trim() + "\"";// "[Subject] > '" + CategoryName.ToLower().Replace("re","").Replace(":","").Trim() +"'";
                //(to:"sajjucode@accsoft.com")
                //string filter = @"(to:\"""+sendTo+"\")";
                //filter= @"[SenderEmailAddress] = """ + sendTo + "\"";
                //filter = @"[ReceivedByName] = """ + sendTo + "\"";
                //filter = @"[ConversationTopic ] = """ + subjectName + "\"";

                //filter = "subject:[New Email]";
                //folderItem = items.Find(filter);
                //folderItem = items.Find(@"[Subject] = """ + CategoryName.ToLower().Replace("re", "").Replace(":", "").Trim() + "\"");

                //while (folderItem != null)
                //{
                //    mailItem = folderItem as MailItem;
                //    if (mailItem != null)
                //    {
                //        //subjectName += "\n" + mailItem.Subject;
                //        ExportSingleMessage(mailItem);
                //    }
                //    folderItem = items.FindNext();
                //}
                //subjectName = " The following e-mail messages were found: " +
                //    subjectName;
                //MessageBox.Show(subjectName);


                Microsoft.Office.Interop.Outlook.MAPIFolder sendFolder = Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail);
                items = sendFolder.Items;
                folderItem = items.Find(filter);
                //folderItem = items.Find(@"[Subject] = """ + CategoryName.ToLower().Replace("re", "").Replace(":", "").Trim() + "\"");

                while (folderItem != null)
                {
                    mailItem = folderItem as MailItem;
                    if (mailItem != null)
                    {
                        //subjectName += "\n" + mailItem.Subject;
                        if (mailItem.ConversationTopic.ToLower().Contains(Subject.ToLower()) || mailItem.Subject.ToLower().Contains(Subject.ToLower()))
                        {
                            ExportAllConversationMessage(mailItem);
                        }
                    }
                    return;
                }
            }
            catch (Exception ex)
      { MessageBox.Show(ex.Message); }
    }


        #endregion
    }
}
