using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace SajjuCode.OutlookAddIns.Base
{
	public class DealManager
	{
		private DateTime DealFile_LastModified;
		private List<Deal> deals = new List<Deal>();
		public string SourceCSV;

		public DealManager() { }
		public DealManager(string csv_path)
		{
			this.SourceCSV = csv_path;
		}

		public void LoadDeals(bool force_reload=false)
		{
			try
			{
				if(force_reload || IsDealFileModified())
				{
					deals = new List<Deal>();
					if (File.Exists(this.SourceCSV))
					{
						this.DealFile_LastModified = File.GetLastWriteTime(this.SourceCSV);
						var all_lines = File.ReadAllLines(this.SourceCSV);
						var skip = false;

						if (all_lines.Length > 1)
						{
							foreach (string line in all_lines)
							{
								//SKIP HEADER
								if (!skip)
								{
									skip = true;
									continue;
								}
								//valid value
								//DEAL INDEX;NAME;VISIBLE;SECTION
								//1;NAME;1;1
								//1;NAME;ADDNAME;3RDNAME;1;1
								var arr = line.Split(';');
								if (arr.Length >= 4)
								{
									var deal = new Deal();
									if (int.TryParse(arr[0].Trim(), out deal.Index))
									{
										deal.Visible = !(arr[arr.Length - 2].Trim() == "0" ||
												arr[arr.Length - 2].Trim().ToLower() == "no" ||
														arr[arr.Length - 2].Trim().ToLower() == "false");

										deal.Section = arr[arr.Length - 1].Trim();
										deal.Name = "";

										for (int j = 1; j < arr.Length - 2; j++)
											deal.Name += arr[j].Trim();

										deal.Name = RemoveSpecialCharacters(deal.Name);
										deals.Add(deal);
									}
								}
							}
						}
					}
				}
			}
			catch(Exception ex) {}			
		}

		public bool IsDealFileModified()
		{
			if(this.DealFile_LastModified!=null && this.SourceCSV != null)
			{
				if (File.GetLastWriteTime(this.SourceCSV) == this.DealFile_LastModified)
				{
					return false;
				}
			}
			return true;
		}

		public bool DealExist(string cleaned_dealname)
		{
			bool ret = false;

			if (this.deals != null)
			{
				if(deals.Where(d=> d.Name.ToLower().Trim() == cleaned_dealname.Trim().ToLower()).Count() > 0)
					ret = true;
			}
			return ret;
		}

		public List<Deal> GetMatchDeal(string dealname)
		{
			List<Deal> match = new List<Deal>();
			dealname = RemoveSpecialCharacters(dealname);

			if (!string.IsNullOrEmpty(dealname))
				match = deals.Where(d => d.Name.ToLower().Trim() == dealname.Trim().ToLower()).ToList();

			return match;
		}

		public List<Deal> GetVisibleDeals()
		{
			List<Deal> ret = new List<Deal>();
			if (this.deals != null)
				ret = deals.Where(d => d.Visible).ToList();

			return ret;
		}

		public int GetDealCount()
		{
			int count = 0;
			if (deals != null)
				count = deals.Count();

			return count;
		}

		public Deal AppendNewDeal(string deal_name)
		{
			Deal d = new Deal();
			bool ok = false;
			StringBuilder append = new StringBuilder();
			
			try
			{
				if (string.IsNullOrEmpty(deal_name)) 
					return null;

				LoadDeals();//refresh list

				//READY FOR ADDING
				d.Index = 1;
				d.Visible = true;
				d.Section = "1";
				d.Name = RemoveSpecialCharacters(deal_name.Trim());

				if (!File.Exists(SourceCSV))
				{
					if (!Directory.Exists(Path.GetDirectoryName(SourceCSV)))
						Directory.CreateDirectory(Path.GetDirectoryName(SourceCSV));

					append.AppendLine("Index;Deal name;Visible;Section");
					append.Append($"1;{deal_name};1;1");
					ok = true;

					File.WriteAllText(SourceCSV,append.ToString());
					this.DealFile_LastModified = File.GetLastWriteTime(this.SourceCSV);
				}
				else
				{
					if(GetMatchDeal(deal_name).Count == 0)
					{
						int last_number = 1;
						if(this.deals.Count > 0)
							last_number = this.deals[deals.Count - 1].Index + 1;

						append.Append($"{Environment.NewLine}{last_number.ToString()};{deal_name};1;1");

						d.Index = last_number;
						this.deals.Add(d);
						File.AppendAllText(SourceCSV, append.ToString());
						this.DealFile_LastModified = File.GetLastWriteTime(this.SourceCSV);
						ok = true;
					}
				}

				if (ok)
					this.deals.Add(d);
				else
					d = null;
			}
			catch (Exception e)
			{
				d = null;
			}
			return d;
		}

		public string RemoveSpecialCharacters(string myInput)
		{
			try
			{
				string myOutPut = myInput;
				if (string.IsNullOrEmpty(myOutPut)) return "";
				// if (!string.IsNullOrEmpty(mySelectedDeals) && mySelectedDeals.ToLower().Contains(DealText.ToLower().Replace(",", "_").Replace("\""," ").Trim()))

				myOutPut = myOutPut.Replace("&", " ");
				myOutPut = myOutPut.Replace(",", "_");
				myOutPut = myOutPut.Replace("\"", " ");
				//myOutPut = myOutPut.Replace(";", " ");

				return myOutPut;
			}
			catch (Exception ex)
			{
				return myInput;
			}
		}

	}
}
