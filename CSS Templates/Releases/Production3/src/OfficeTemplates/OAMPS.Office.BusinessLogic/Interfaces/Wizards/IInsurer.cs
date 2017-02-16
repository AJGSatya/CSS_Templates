
namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IInsurer
    {
        string Title { get; set; }
        string Id { get; set; }
        string Category { get; set; }
        bool Selected { get; set; }
        decimal Percent { get; set; }
    }
}
