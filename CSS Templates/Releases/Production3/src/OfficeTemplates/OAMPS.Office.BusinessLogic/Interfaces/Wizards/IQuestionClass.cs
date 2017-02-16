using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IQuestionClass
    {
        string Title { get; set; }
        string TopCategory { get; set; }
        string SubCategory { get; set; }
        string Url { get; set; }
        string Id { get; set; }
    }
}
