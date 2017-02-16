using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IInsuranceManual : IBaseTemplate
    {
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

        bool PopulateClientProfile { get; set; }
        List<IPolicyClass> SelectedPolicyClasses { get; set; }
    }
}