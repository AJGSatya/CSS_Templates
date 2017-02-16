namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IQuoteSlipSchedules
    {
        string Title { get; set; }
        string TopCategory { get; set; }
        string SubCategory { get; set; }
        string Url { get; set; }
        string Id { get; set; }
        string[] LinkedQuestionUrl { get; set; }
        string[] LinkedQuestionId { get; set; }
    }
}