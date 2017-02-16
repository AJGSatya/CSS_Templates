using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Template
{
    public class GenericLetter : BaseTemplate, IGenericLetter
    {
        public string Addressee { get; set; }
        public string Salutation { get; set; }
        public string ClientAddress { get; set; }
        public string ClientAddressLine2 { get; set; }
        public string ClientAddressLine3 { get; set; }               
        public string OAMPSBranchAddress { get; set; }
        
        public string OAMPSBranchAddressLine2 { get; set; }
        public string OAMPSPostalAddress { get; set; }
        public string OAMPSPostalAddressLine2 { get; set; }
        public string Fax { get; set; }

        public string DatePrepared { get; set; }
        public string Reference { get; set; }
        public string Subject { get; set; }

        public bool IsAdviceWarningSelected { get; set; }
        public bool IsRisksSelected { get; set; }
        public bool IsPrivacySelected { get; set; }
        public bool IsSatutorySelected { get; set; }
        public bool IsFSGSelected { get; set; }
        public string PDSDescVersion { get; set; }

        public bool IsPrePrintSelected { get; set; }

        public IPolicyClass PolicyType { get; set; }
    }
}
