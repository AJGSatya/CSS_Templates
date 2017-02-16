using OAMPS.Office.BusinessLogic.Interfaces.Template;

namespace OAMPS.Office.BusinessLogic.Models.Template
{
    public class Agenda : BaseTemplate, IAgenda
    {
        public string AgendaDate { get; set; }
        public string AgendaTimeFrom { get; set; }
        public string AgendaTimeTo { get; set; }
        public string Location { get; set; }
        public string Subject { get; set; }
    }
}