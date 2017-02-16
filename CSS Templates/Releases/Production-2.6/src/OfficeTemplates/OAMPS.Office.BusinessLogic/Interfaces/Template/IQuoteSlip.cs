using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IQuoteSlip : IBaseTemplate
    {
        
        string PeriodOfInsuranceFrom { get; set; }
        string PeriodOfInsuranceTo { get; set; }
        
        string OAMPSPostalAddress { get; set; }
        string OAMPSPostalAddressLine2 { get; set; }

        string AssistantExecutiveName { get; set; }
        string AssistantExecutiveTitle { get; set; }
        string AssistantExecutivePhone { get; set; }
        string AssistantExecutiveEmail { get; set; }
        string AssistantExecDepartment { get; set; }

      
        string DatePrepared { get; set; }

        bool PopulateClaimMadeWarning { get; set; }
        bool PopulateApprovalForm { get; set; }

        List<IPolicyClass> SelectedDocumentFragments { get; set; }
    }
}