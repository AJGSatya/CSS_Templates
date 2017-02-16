using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface ISummaryOfDiscussions : IBaseTemplate
    {
        string ClientCode { get; set; }
        string ClientContactName { get; set; }

        string DateDiscussion { get; set; }
        string TimeDiscussion { get; set; }


        string IsDiscussedByPhone { get; set; }
        string IsDiscussedInPerson { get; set; }
        string IsOther { get; set; }

        string IsDiscussedWithCustomer { get; set; }
        string IsDiscussedWithCaller { get; set; }
        string IsDiscussedWithUnderWriter { get; set; }
    }
}
