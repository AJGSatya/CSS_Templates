using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class QuestionClass: IQuestionClass
    {
        public string Title { get; set; }
        public string TopCategory { get; set; }
        public string SubCategory { get; set; }
        public string Url { get; set; }
        public string Id { get; set; }
    }
}
