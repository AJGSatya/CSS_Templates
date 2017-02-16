using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class PolicyClass : IPolicyClass
    {

        public string TitleNoWhiteSpace { get; set; }
        public string Title { get; set; }
        public string CurrentInsurer { get; set; }
        public string RecommendedInsurer { get; set; }
        public string MajorClass { get; set; }
        public string Url { get; set; }
        public string Id { get; set; }
        public string FragmentPolicyUrl { get; set; }
        public string CurrentInsurerId { get; set; }
        public string RecommendedInsurerId { get; set; }
        public int Order { get; set; }
    }
}
