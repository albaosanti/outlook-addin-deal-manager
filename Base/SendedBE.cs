using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SajjuCode.OutlookAddIns.Base
{
    public class SendedBE
    {
        public string Topic { get; set; }
        public string ConversationID { get; set; }
        public string SendTo { get; set; } 
        public Boolean Sended { get; set; }
    }
}
