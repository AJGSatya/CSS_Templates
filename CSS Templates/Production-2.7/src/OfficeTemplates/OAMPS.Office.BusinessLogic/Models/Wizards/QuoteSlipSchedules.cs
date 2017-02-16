using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class QuoteSlipSchedules : IQuoteSlipSchedules
    {
        public string Title { get; set; }
        public string TopCategory { get; set; }
        public string SubCategory { get; set; }
        public string Url { get; set; }
        public string Id { get; set; }
        public string[] LinkedQuestionUrl { get; set; }
        public string[] LinkedQuestionId { get; set; }
    }
}