using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Template;

namespace OAMPS.Office.BusinessLogic.Models.Template
{
    public class BaseTemplate : IBaseTemplate
    {
        public string ClientName { get; set; }
        public string DocumentTitle { get; set; }
        public string DocumentSubTitle { get; set; }
        public string CoverPageImageUrl { get; set; }
        public string LogoImageUrl { get; set; }
        public string CoverPageTitle { get; set; }
        public string LogoTitle { get; set; }
        public string OAMPSCompanyName { get; set; }
        public string OAMPSAbnNumber { get; set; }
        public string OAMPSBranchPhone { get; set; }
        public string OAMPSAfsl { get; set; }
        public string WebSite { get; set; }

        public string ExecutiveName { get; set; }
        public string ExecutiveTitle { get; set; }
        public string ExecutivePhone { get; set; }
        public string ExecutiveMobile { get; set; }
        public string ExecutiveEmail { get; set; }
        public string ExecutiveDepartment { get; set; }
        public string LongBrandingDescription { get; set; }
    }
}
