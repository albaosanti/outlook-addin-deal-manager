using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SajjuCode.OutlookAddIns.Base
{
    public class SettingsClassFunctions
    {
        public static string OfferCSVPath = "";
        public static string OfferCSVFolder= "";
        public static string MessageSaveFolder = "";
        public static string getSettingFile
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "nwtDeals.txt");
            }
        }

        public static bool isDealFileExists
        {
            get
            {
                return File.Exists(getSettingFile);
            }
        }

        public static bool ReadSettingFile()
        {
            bool ok = false;
            try
            {
                if (isDealFileExists==true)
                {
                    string mySettingData = File.ReadAllText(getSettingFile);
                    if (!string.IsNullOrEmpty(mySettingData))
                    {
                        int i = 0;
                        foreach(string myS in mySettingData.Split(';'))
                        {
                            if (string.IsNullOrEmpty(myS))
                            {
                                continue;
                            }

                            if (i==0)
                            {
                                OfferCSVPath = myS;
                                if (File.Exists(myS))
                                {
                                    FileInfo myFileI = new FileInfo(myS);
                                    OfferCSVFolder= myFileI.DirectoryName;
                                }
                            }
                            else if (i==1)
                            {
                                MessageSaveFolder = myS;
                            }

                            i++;
                        }
                        if(!string.IsNullOrEmpty(OfferCSVFolder) && !string.IsNullOrEmpty(MessageSaveFolder))
                        {
                          ok = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {}
            return ok;
        }
    }
}
