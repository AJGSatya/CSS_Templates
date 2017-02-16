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
                string cs = "CSS Templates";

                var elog = new EventLog();
                if (!EventLog.SourceExists(cs))
                {
                    try
                    {
                        EventLog.CreateEventSource(cs, cs);
                    }
                    catch (Exception)
                    {
                        cs = "Application";
                    }
                }
                elog.Source = cs;
                elog.EnableRaisingEvents = true;
                elog.WriteEntry(message, GetLogType(type));
            }
            catch (Exception)
            {
                //todo: hmmm decide what we do here 
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