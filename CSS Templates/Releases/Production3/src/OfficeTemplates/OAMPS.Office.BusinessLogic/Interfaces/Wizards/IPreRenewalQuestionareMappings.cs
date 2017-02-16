using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IPreRenewalQuestionareMappings
    {
        string FragmentUrl { get; set; }
        string Title { get; set; }
        string[] PolicyType { get; set; }
    }
}
