using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class PremiumSummaryTableRow : IPremiumSummaryTableRow
    {
        public string ClassOfInsurance { get; set; }
        public string RecommendedInsurer { get; set; }
        public string BasePremium { get; set; }
        public string FireServicesLevies { get; set; }
        public string TotalGst { get; set; }
        public string PolicyUnderwriterGst { get; set; }
        public string StampDuty { get; set; }
        public string OtherTaxesCharges { get; set; }
        public string BrokerFeePerPolicyClass { get; set; }
        public string PeriodOfInsuranceFrom { get; set; }
        public string PeriodOfInsuranceTo { get; set; }
        public string Comments { get; set; }
    }
}