using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    internal interface IGenericLetter : IBaseTemplate
    {
        string Addressee { get; set; }

        string Salutation { get; set; }
        string ClientAddress { get; set; }
        string ClientAddressLine2 { get; set; }
        string ClientAddressLine3 { get; set; }

        string OAMPSPostalAddress { get; set; }
        string OAMPSPostalAddressLine2 { get; set; }

        string DatePrepared { get; set; }
        string Reference { get; set; }
        string Subject { get; set; }
        bool IsAdviceWarningSelected { get; set; }
        bool IsRisksSelected { get; set; }
        bool IsPrivacySelected { get; set; }
        bool IsSatutorySelected { get; set; }
        bool IsFsgSelected { get; set; }
        bool IsPrePrintSelected { get; set; }
        string PdsDescVersion { get; set; }

        IPolicyClass PolicyType { get; set; }
    }
}