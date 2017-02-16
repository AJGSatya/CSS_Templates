using System;
using System.Diagnostics;
using OAMPS.Office.BusinessLogic.Interfaces.Logging;
using Type = OAMPS.Office.BusinessLogic.Interfaces.Logging.Type;

namespace OAMPS.Office.BusinessLogic.Loggers
{
    public class EventViewerLogger : IEventViewerLogger
    {
        public void Log(string message, Type type)
        {
            try
            {
                var cs = "CSS Templates";

                var elog = new EventLog();
                if (!EventLog.SourceExists(cs))
                {
                    try
                    {
                        EventLog.CreateEventSource(cs, cs);
                    }
                    catch (Exception ex)
                    {
                        cs = "Application";
                        Debug.WriteLine("[CSSTemplates] Could not create eventlog source: {0}, ex: {1}", cs, ex);
                    }
                }
                elog.Source = cs;
                elog.EnableRaisingEvents = true;
                elog.WriteEntry(message, GetLogType(type));

                Debug.WriteLine("[CSSTemplates] Successfully logged to {0} : {1}", cs, message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[CSSTemplates] Could not write to eventlog {0}, error {1} ", message, ex);
            }
        }

        private EventLogEntryType GetLogType(Type type)
        {
            switch (type)
            {
                case Type.Error:
                    return EventLogEntryType.Error;

                case Type.Debug:
                    return EventLogEntryType.Information;

                case Type.Information:
                    return EventLogEntryType.Information;

                case Type.Warning:
                    return EventLogEntryType.Warning;

                default:
                    return EventLogEntryType.Information;
            }
        }
    }
}