using System.Data.Entity;

namespace OAMPS.Office.Word.Models
{
    public class UsageLoggingContext : DbContext
    {
        public UsageLoggingContext() : base("CSSTemplatesConnectionString")
        {
        }


        public DbSet<UsageLog> UsageLog { get; set; }
    }
}