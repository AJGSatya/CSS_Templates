using OAMPS.Office.BusinessLogic.Interfaces.Logging;
using OAMPS.Office.Word.Models;

namespace OAMPS.Office.Word.Helpers
{
    public class DbUsageLogger : IUsageLogger
    {
        private readonly UsageLog _usageLog;

        public DbUsageLogger(UsageLog usageLog)
        {
            _usageLog = usageLog;
        }

        public void LogUsage()
        {
            using (var db = new UsageLoggingContext())
            {
                db.UsageLog.Add(_usageLog);
                db.SaveChanges();
            }
        }
    }
}