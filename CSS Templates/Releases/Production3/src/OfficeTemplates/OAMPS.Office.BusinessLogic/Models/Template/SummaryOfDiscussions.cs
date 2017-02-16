using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Template;

namespace OAMPS.Office.BusinessLogic.Models.Template
{
    public class SummaryOfDiscussions : BaseTemplate, ISummaryOfDiscussions
    {
        public string ClientCode { get; set; }
        public string ClientContactName { get; set; }

        public string DateDiscussion { get; set; }
        public string TimeDiscussion { get; set; }

        public string IsDiscussedByPhone { get; set; }
        public string IsDiscussedInPerson { get; set; }
        public string IsOther { get; set; }
        public string IsDiscussedWithCustomer { get; set; }
        public string IsDiscussedWithCaller { get; set; }
        public string IsDiscussedWithUnderWriter { get; set; }
    }
}
