using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IPrenewalQuestionare : IBaseTemplate
    {
        string ClientCommonName { get; set; }
        
        string OAMPSBranchAddress { get; set; }
        string OAMPSBranchAddressLine2 { get; set; }
        string PeriodOfInsuranceFrom { get; set; }
        string PeriodOfInsuranceTo { get; set; }

        string AssistantExecutiveName { get; set; }
        string AssistantExecutiveTitle { get; set; }
        string AssistantExecutivePhone { get; set; }
        string AssistantExecutiveEmail { get; set; }
        string AssistantExecDepartment { get; set; }

        string ClaimsExecutiveName { get; set; }
        string ClaimsExecutiveTitle { get; set; }
        string ClaimsExecutivePhone { get; set; }
        string ClaimsExecutiveEmail { get; set; }
        string ClaimsExecDepartment { get; set; }

        string OtherContactName { get; set; }
        string OtherContactTitle { get; set; }
        string OtherContactPhone { get; set; }
        string OtherContactEmail { get; set; }
        string OtherExecDepartment { get; set; }

        string DatePrepared { get; set; }
        
        bool PopulateClaimMadeWarning { get; set; }
        bool PopulateApprovalForm { get; set; }

        List<IPolicyClass> SelectedDocumentFragments { get; set; } 
    }
}
