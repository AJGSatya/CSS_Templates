using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OAMPS.Office.BusinessLogic.Interfaces.Wizards;


namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IClientDiscovery
    {
        string ClientName { get; set; }
        string ClientContactName { get; set; }
        string ClientiBais { get; set; }        
        string DiscussionDate { get; set; }

        string OAMPSBranchAddress { get; set; }
        string OAMPSBranchAddressLine2 { get; set; }

        string ExecutiveName { get; set; }
        string ExecutiveTitle { get; set; }
        string ExecutivePhone { get; set; }
        string ExecutiveMobile { get; set; }
        string ExecutiveEmail { get; set; }

        List<IQuestionClass> Questions { get; set; } 

        
    }
}
