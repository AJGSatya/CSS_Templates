using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class ManualClaimsProcedure : IManualClaimsProcedure
    {
        public string Title { get; set; }
        public string PolicyClass { get; set; }
        public string Url { get; set; }
        public string Id { get; set; }
    }
}