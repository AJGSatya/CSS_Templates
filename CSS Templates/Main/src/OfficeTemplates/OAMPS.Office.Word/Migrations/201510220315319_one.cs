using System.Data.Entity.Migrations;

namespace OAMPS.Office.Word.Migrations
{
    public partial class One : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.UsageLogs",
                c => new
                {
                    UsageLogId = c.Int(false, true),
                    TrackingType = c.String(),
                    Report = c.String(),
                    AccountExecutive = c.String(),
                    UserDepartment = c.String(),
                    UserOffice = c.String(),
                    ClientName = c.String(),
                    Segment = c.String(),
                    WholesaleOrRetail = c.String(),
                    CaptureDate = c.String(),
                    CaptureTime = c.String(),
                    VersionNumber = c.String(),
                    UserName = c.String()
                })
                .PrimaryKey(t => t.UsageLogId);
        }

        public override void Down()
        {
            DropTable("dbo.UsageLogs");
        }
    }
}