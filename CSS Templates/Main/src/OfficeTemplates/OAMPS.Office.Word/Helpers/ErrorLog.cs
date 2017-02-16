using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OAMPS.Office.Word.Properties;
using Serilog;

namespace OAMPS.Office.Word.Helpers
{
    public static class ErrorLog
    {        
        private static readonly ILogger Logger = new LoggerConfiguration()
            .MinimumLevel.Error()
            .WriteTo.RollingFile(Environment.ExpandEnvironmentVariables(Settings.Default.ErrorLogPath) + "\\error-{Date}.txt")
            .CreateLogger();

        private static readonly ILogger TraceLogger = new LoggerConfiguration()
            .WriteTo.RollingFile(Environment.ExpandEnvironmentVariables(Settings.Default.ErrorLogPath) + "\\trace-{Date}.txt")
            .CreateLogger();

        public static void Error(Exception ex, string message, params object[] args)
        {
            Logger.Error(ex, message, args);
        }

        public static void Error(string message, params object[] args)
        {
            Logger.Error(message, args);
        }

        public static void TraceLog(string message, params object[] args)
        {
            TraceLogger.Information(message, args);
        }
    }
}
