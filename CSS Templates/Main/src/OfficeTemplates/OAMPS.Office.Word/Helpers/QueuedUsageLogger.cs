using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OAMPS.Office.Word.Models;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Helpers
{
    public sealed class QueuedUsageLogger
    {
        private static volatile QueuedUsageLogger _instance;
        private static object _syncRoot = new object();
        private static object _lock = new object();

        private ConcurrentQueue<UsageLog> _queue;
        private readonly string _queueFilePath;

        private bool _transactionActive;
        private bool _failedToWrite;

        private QueuedUsageLogger()
        {
            _queue = new ConcurrentQueue<UsageLog>();
            _queueFilePath = Environment.ExpandEnvironmentVariables(Settings.Default.UsageLogQueue);

            _transactionActive = false;
            _failedToWrite = false;

            if (File.Exists(_queueFilePath))
            {
                var data = File.ReadAllText(_queueFilePath);
                var list = JsonConvert.DeserializeObject<List<UsageLog>>(data);
                foreach (var entry in list)
                {
                    _queue.Enqueue(entry);
                }
            }
        }

        public static QueuedUsageLogger Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_syncRoot)
                    {
                        if (_instance == null)
                            _instance = new QueuedUsageLogger();
                    }
                }
                return _instance;
            }
        }

        public void LogUsage(UsageLog log)
        {
            _queue.Enqueue(log);

            // only one thread ever to write
            if (_transactionActive)
                return;
            lock(_lock)
            {
                if (_transactionActive)
                    return;
                _transactionActive = true;
            }

            if (_failedToWrite)
            {
                WriteUsageToDisk();
            }
            else
            {
                if (!WriteUsageToDataBase())
                {
                    _failedToWrite = true;
                    WriteUsageToDisk();
                }
            }
            _transactionActive = false;
        }

        private void WriteUsageToDisk()
        {
            var output = JsonConvert.SerializeObject(_queue);
            File.WriteAllText(_queueFilePath, output);
        }

        private bool WriteUsageToDataBase()
        {
            UsageLog logEntry = null;
            try
            {
                while (!_queue.IsEmpty)
                {
                    var hasItem = _queue.TryDequeue(out logEntry);
                    if (hasItem && logEntry != null)
                    {
                        var usageLogger = new DbUsageLogger(logEntry);
                        usageLogger.LogUsage();
                    }
                }
                return true;
            }
            catch (ThreadAbortException)
            {
                if (logEntry != null)
                    _queue.Enqueue(logEntry);
                _failedToWrite = true;
                return false;
            }
            catch (Exception ex)
            {
                if (logEntry != null)
                    _queue.Enqueue(logEntry);
                _failedToWrite = true;

                ErrorLog.Error(ex, "QueuedUsageLogger.WriteUsageLog, items in queue: {0}", _queue.Count);
                return false;
            }
            
        }
    }
}
