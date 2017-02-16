namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IAgenda : IBaseTemplate
    {
        string AgendaDate { get; set; }
        string AgendaTimeFrom { get; set; }
        string AgendaTimeTo { get; set; }
        string Location { get; set; }
        string Subject { get; set; }
    }
}