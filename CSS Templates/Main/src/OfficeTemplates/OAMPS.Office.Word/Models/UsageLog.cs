namespace OAMPS.Office.Word.Models
{
    public class UsageLog
    {
        public int UsageLogId { get; set; }
        public string TrackingType { get; set; }
        public string Report { get; set; }
        public string AccountExecutive { get; set; }
        public string UserDepartment { get; set; }
        public string UserOffice { get; set; }
        public string ClientName { get; set; }
        public string Segment { get; set; }
        public string WholesaleOrRetail { get; set; }
        public string CaptureDate { get; set; }
        public string CaptureTime { get; set; }
        public string VersionNumber { get; set; }
        public string UserName { get; set; }
    }
}