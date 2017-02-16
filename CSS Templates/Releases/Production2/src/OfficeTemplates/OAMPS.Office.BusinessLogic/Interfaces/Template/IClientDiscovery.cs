using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OAMPS.Office.BusinessLogic.Interfaces.Wizards;


namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IClientDiscovery : IBaseTemplate
    {
        string ClientContactName { get; set; }
        string ClientiBais { get; set; }        
        string DiscussionDate { get; set; }

        string OAMPSBranchAddress { get; set; }
        string OAMPSBranchAddressLine2 { get; set; }


        List<IQuestionClass> Questions { get; set; } 

        
    }
}
