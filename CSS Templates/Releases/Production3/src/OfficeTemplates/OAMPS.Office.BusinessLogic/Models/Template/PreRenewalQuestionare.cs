using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Template
{
    public class PreRenewalQuestionare : BaseTemplate, IPrenewalQuestionare
    {        
        public string ClientCommonName { get; set; }
        public string OAMPSBranchAddress { get; set; }
        public string OAMPSBranchAddressLine2 { get; set; }
        public string PeriodOfInsuranceFrom { get; set; }
        public string PeriodOfInsuranceTo { get; set; }
        public string AssistantExecutiveName { get; set; }
        public string AssistantExecutiveTitle { get; set; }
        public string AssistantExecutivePhone { get; set; }
        public string AssistantExecutiveEmail { get; set; }
        public string AssistantExecDepartment { get; set; }
        public string ClaimsExecutiveName { get; set; }
        public string ClaimsExecutiveTitle { get; set; }
        public string ClaimsExecutivePhone { get; set; }
        public string ClaimsExecutiveEmail { get; set; }
        public string ClaimsExecDepartment { get; set; }
        public string OtherContactName { get; set; }
        public string OtherContactTitle { get; set; }
        public string OtherContactPhone { get; set; }
        public string OtherContactEmail { get; set; }
        public string OtherExecDepartment { get; set; }
        public string DatePrepared { get; set; }
        
        public bool PopulateClaimMadeWarning { get; set; }
        public bool PopulateApprovalForm { get; set; }


        public List<IPolicyClass> SelectedDocumentFragments { get; set; }
    }
}
