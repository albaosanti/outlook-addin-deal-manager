using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SajjuCode.OutlookAddIns
{
    public partial class ThisAddIn
    {
        private void AddACategory()
        {
            Outlook.Categories categories =Application.Session.Categories;
            if (!CategoryExists("Sajjucode2"))
            {
                Outlook.Category category = categories.Add("Sajjucode2",Outlook.OlCategoryColor.olCategoryColorDarkBlue);
            }

            //CreateTextAndCategoryRule();
            //AssignCategories();
        }

        private void RemoveCategory(string CategoryName="")
        {
            try
            {
                Outlook.Categories categories = Application.Session.Categories;
                List<int> listIndex = new List<int>();
                if (categories !=null && categories.Count>0)
                {
                   for (int i=1; i<categories.Count;i++)
                    {
                        var myC = categories[i];

                        if (myC.Name.Length>1)
                        {
                            categories.Remove(myC.Name);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {

                
            }
        }

        private bool CategoryExists(string categoryName)
        {
            try
            {
                Outlook.Category category =Application.Session.Categories[categoryName];
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

        #region "Assignment"
        private void CreateTextAndCategoryRule()
        {
            if (!CategoryExists("Sajjucode2"))
            {
                Application.Session.Categories.Add("Office", Type.Missing, Type.Missing);
            }
            if (!CategoryExists("Sajjucode2"))
            {
                Application.Session.Categories.Add("Sajjucode2", Type.Missing, Type.Missing);
            }
            Outlook.Rules rules =
                Application.Session.DefaultStore.GetRules();
            Outlook.Rule textRule =rules.Create("Demo Text and Category Rule Sajjucode",Outlook.OlRuleType.olRuleReceive);
            Object[] textCondition ={ "Major update from Message center", "Office", "Outlook", "2007" };
            Object[] categoryAction ={"Sajjucode2", "Office", "Outlook" };
            textRule.Conditions.BodyOrSubject.Text =textCondition;
            textRule.Conditions.BodyOrSubject.Enabled = true;
            textRule.Actions.AssignToCategory.Categories =categoryAction;
            textRule.Actions.AssignToCategory.Enabled = true;
            
            rules.Save(true);
        }

        private void AssignCategories()
        {
            string filter = "@SQL=" + "\"" + "urn:schemas:httpmail:subject"
                + "\"" + " ci_phrasematch 'ISV'";

            var myII = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items;

            //Outlook.Items items =Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items.Restrict(filter);
            for (int i = 1; i <= 5; i++)
            {
                OutlookItem item = new OutlookItem(myII[i]);
                string existingCategories = item.Categories;
                if (String.IsNullOrEmpty(existingCategories))
                {
                    item.Categories = "ISV";
                }
                else
                {
                    if (item.Categories.Contains("ISV") == false)
                    {
                        item.Categories = existingCategories + ", ISV";
                    }
                }
                item.Save();
            }
        }

        #endregion
    }
}
