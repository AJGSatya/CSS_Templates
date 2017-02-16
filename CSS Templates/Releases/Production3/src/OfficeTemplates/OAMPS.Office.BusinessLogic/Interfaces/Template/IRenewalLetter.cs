using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    interface IRenewalLetter : IBaseTemplate
    {
        string Addressee { get; set; }
        string Salutation { get; set; }
        string ClientAddress { get; set; }
        string ClientAddressLine2 { get; set; }
        string ClientAddressLine3 { get; set; }        
                
        string OAMPSBranchAddress { get; set; }
        string OAMPSBranchAddressLine2 { get; set; }

        string OAMPSPostalAddress { get; set; }
        string OAMPSPostalAddressLine2 { get; set; }
        string Fax { get; set; }
        
                
        string DatePrepared { get; set; }
        string DatePolicyExpiry { get; set; }
        string Reference { get; set; }
        string PaymentDate { get; set; }

        bool IsNewClientSelected { get; set; }
        bool IsContactSelected { get; set; }
        bool IsFundingSelected { get; set; }

        bool IsAdviceWarningSelected { get; set; }
        bool IsRisksSelected { get; set; }
        bool IsPrivacySelected { get; set; }
        bool IsSatutorySelected { get; set; }
        bool IsFSGSelected { get; set; }
        bool IsGAWSelected { get; set; }

        string PDSDescVersion { get; set; }

         List<IPolicyClass> Policies { get; set; }
        
    }
}
