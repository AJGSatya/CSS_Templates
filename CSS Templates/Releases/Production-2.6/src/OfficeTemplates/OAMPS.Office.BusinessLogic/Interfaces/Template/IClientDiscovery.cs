using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IClientDiscovery : IBaseTemplate
    {
        string ClientContactName { get; set; }
        string ClientiBais { get; set; }
        string DiscussionDate { get; set; }

        List<IQuestionClass> Questions { get; set; }

        string DatePrepared { get; set; }
    }
}