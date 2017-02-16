namespace OAMPS.Office.BusinessLogic.Interfaces.Template
{
    public interface IBaseTemplate
    {
        string ClientName { get; set; }
        string ClientCommonName { get; set; }
        string OAMPSBranchAddress { get; set; }
        string OAMPSBranchAddressLine2 { get; set; }
        string Fax { get; set; }

        string DocumentTitle { get; set; }
        string DocumentSubTitle { get; set; }
        string CoverPageImageUrl { get; set; }
        string LogoImageUrl { get; set; }
        string CoverPageTitle { get; set; }
        string LogoTitle { get; set; }
        string OAMPSCompanyName { get; set; }
        string OAMPSAbnNumber { get; set; }
        string OAMPSBranchPhone { get; set; }
        string OAMPSAfsl { get; set; }
        string WebSite { get; set; }

        string ExecutiveName { get; set; }
        string ExecutiveTitle { get; set; }
        string ExecutivePhone { get; set; }
        string ExecutiveMobile { get; set; }
        string ExecutiveEmail { get; set; }
        string ExecutiveDepartment { get; set; }

        string LongBrandingDescription { get; set; }
    }
}