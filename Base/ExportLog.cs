using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace SajjuCode.OutlookAddIns.Base
{
	public class ExportLog
	{
		public string LogPath;

		public void WriteLog(string message,bool err,bool prompt = false)
		{
			string log_filename;
			try{
				if(!Directory.Exists(LogPath))
					Directory.CreateDirectory(LogPath);

				log_filename = DateTime.Now.ToString("yyyy-MM-dd");
				if (err)
					log_filename += ".err";
				else
					log_filename += ".log";

				StringBuilder msg = new StringBuilder();
				msg.AppendLine($"{DateTime.Now.ToString()}\t{message}");

				File.AppendAllText(LogPath + "\\" + log_filename, msg.ToString());
			}
			catch (Exception ex)
			{}
			if (prompt) System.Windows.Forms.MessageBox.Show(message,"Outlook AddIn - Log");
		}
	}
}
