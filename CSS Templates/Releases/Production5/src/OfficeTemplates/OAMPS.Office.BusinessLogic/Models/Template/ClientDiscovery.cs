using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Template
{
    public class ClientDiscovery : BaseTemplate, IClientDiscovery
    {
        public string ClientContactName { get; set; }
        public string ClientiBais { get; set; }
        public string DiscussionDate { get; set; }

        public List<IQuestionClass> Questions { get; set; }
        public string DatePrepared { get; set; }
    }
}