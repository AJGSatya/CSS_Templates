﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    interface IRenewalLetter
    {
        string Addressee { get; set; }
        string ClientName { get; set; }
        string Salutation { get; set; }
        string ClientAddress { get; set; }
        string ClientAddressLine2 { get; set; }        
                
        string OAMPSBranchAddress { get; set; }
        string OAMPSBranchAddressLine2 { get; set; }

        string OAMPSPostalAddress { get; set; }
        string OAMPSPostalAddressLine2 { get; set; }
        string Fax { get; set; }
        string ExecutiveName { get; set; }
        string ExecutiveTitle { get; set; }
        string ExecutivePhone { get; set; }
        string ExecutiveMobile { get; set; }
        string ExecutiveEmail { get; set; }
                
        string DatePrepared { get; set; }
        string DatePolicyExpiry { get; set; }
        string Reference { get; set; }
        string PaymentDate { get; set; }

        bool IsNewClientSelected { get; set; }
        bool IsContactSelected { get; set; }
        bool IsFundingSelected { get; set; }

        bool IsWarningSelected { get; set; }
        bool IsRisksSelected { get; set; }
        bool IsPrivacySelected { get; set; }
        bool IsSatutorySelected { get; set; }
        bool IsFSGSelected { get; set; }

        string PDSDescVersion { get; set; }
 
        IPolicyClass PolicyType { get; set; }
        
    }
}
