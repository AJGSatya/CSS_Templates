using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Template
{
    public class QuoteSlip : BaseTemplate, IQuoteSlip
    {
        public string DateRequiredBy { get; set; }
        public string OAMPSPostalAddress { get; set; }
        public string OAMPSPostalAddressLine2 { get; set; }

        public string PeriodOfInsuranceFrom { get; set; }
        public string PeriodOfInsuranceTo { get; set; }
        public string AssistantExecutiveName { get; set; }
        public string AssistantExecutiveTitle { get; set; }
        public string AssistantExecutivePhone { get; set; }
        public string AssistantExecutiveEmail { get; set; }
        public string AssistantExecDepartment { get; set; }

        public string DatePrepared { get; set; }

        public bool PopulateClaimMadeWarning { get; set; }
        public bool PopulateApprovalForm { get; set; }


        public List<IPolicyClass> SelectedDocumentFragments { get; set; }
    }
}