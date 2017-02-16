using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class PreRenewalQuestionareMappings : IPreRenewalQuestionareMappings
    {
        public string FragmentUrl { get; set; }
        public string Title { get; set; }
        public string[] PolicyType { get; set; }
    }
}