using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SajjuCode.OutlookAddIns.Base
{
	public class OfferteConversation
	{
		public string DealIndex { get; set; }
		public string ConversationID { get; set; }
		public Outlook.Conversation conversation { get; set; }
		public OfferteConversation() { }

		public OfferteConversation(string conversationID,Outlook.Conversation conv)
		{
			this.ConversationID = conversationID;
			this.conversation = conv;
		}
	}
}
