
namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IPremiumSummaryTableRow
    {
        string ClassOfInsurance { get; set; }
        string RecommendedInsurer { get; set; }
        string BasePremium { get; set; }
        string FireServicesLevies { get; set; }
        string TotalGST { get; set; }
        string PolicyUnderwriterGst { get; set; }
        string StampDuty { get; set; }
        string OtherTaxesCharges { get; set; }
        string BrokerFeePerPolicyClass { get; set; }
        string PeriodOfInsuranceFrom { get; set; }
        string PeriodOfInsuranceTo { get; set; }
        string Comments { get; set; }
    }
}
