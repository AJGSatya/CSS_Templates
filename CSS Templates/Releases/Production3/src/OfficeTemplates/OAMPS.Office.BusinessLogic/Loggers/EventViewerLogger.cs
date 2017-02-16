using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
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
                    catch (Exception cException)
                    {
                        cs = "Application";
                    }
                }
                elog.Source = cs;
                elog.EnableRaisingEvents = true;
                elog.WriteEntry(message, GetLogType(type));
            }
            catch (Exception logEx)
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
                    break;

                case Type.Debug:
                    return EventLogEntryType.Information;
                    break;

                case Type.Information:
                    return EventLogEntryType.Information;
                    break;

                case Type.Warning:
                    return EventLogEntryType.Warning;
                    break;

                default:
                    return EventLogEntryType.Information;
                    break;
            }
        }
    }
}
