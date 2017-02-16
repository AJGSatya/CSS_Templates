
namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IPolicyClass
    {
        string TitleNoWhiteSpace { get; set; }
        string Title { get; set; }
        string CurrentInsurer { get; set; }
        string RecommendedInsurer { get; set; }
        string CurrentInsurerId { get; set; }
        string RecommendedInsurerId { get; set; }
        string MajorClass { get; set; }
        string Url { get; set; }
        int Order { get; set; }
        string Id { get; set; }
    }
}
