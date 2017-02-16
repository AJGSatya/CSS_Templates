using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSTemplates.TemplateSyncTool.Settings;
using Newtonsoft.Json;
using OAMPS.Office.BusinessLogic.Models;

namespace CSSTemplates.TemplateSyncTool.Synchronise
{
    class Synchroniser
    {
        public void Synchronize(log4net.ILog logger, Arguments args)
        {
            logger.Info($"\n\n*** Synchronisation of fragments started - {DateTime.Now} ***\n\n");
            
            var updateCounter = new UpdateCounter();
            var updateLog = new UpdateLog();
            var updateLogFileName = Path.Combine(args.Directory.FullName, "UpdateLog.txt");
            if (File.Exists(updateLogFileName))
            {
                var data = File.ReadAllText(updateLogFileName);
                updateLog = JsonConvert.DeserializeObject<UpdateLog>(data);
            }

            try
            {
                if (args.Clean)
                {
                    logger.Info($"\nCleaning out target directory {args.Directory.Name}..\n\n");
                    ClearFolder(args.Directory);
                }
                Directory.CreateDirectory(args.Directory.FullName);

                var contextUrl = args.Source;
                logger.Info("\nConnecting to SharePoint online..\n\n");

                using (var spContext = new SharePointClientContext(contextUrl))
                {
                    logger.Info($"Loading fragments from site: {contextUrl}");

                    var syncConf = FragmentSettings.GetConfig();
                    var fileSync = new FileSynchroniser();
                    var listSync = new ListSynchroniser();
                    foreach (Fragment fragment in syncConf.Fragments)
                    {
                        try
                        {
                            if (fragment.Type == "library")
                            {
                                fileSync.Synchronise(spContext, fragment, updateCounter, updateLog, logger, args);
                            }
                            else if (fragment.Type == "list")
                            {
                                listSync.Synchronise(spContext, fragment, updateCounter, updateLog, logger, args);
                            }
                            else
                            {
                                logger.Warn($"\n\nWarning: Unsupported synchronisation settig found; type:{fragment.Type}, source:{fragment.Source}, destination:{fragment.Destination}\n\n");
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Error($"Failed synchronising {fragment.Source}", ex);
                            if (args.BreakOnError)
                                throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("An exception was thrown", ex);
                if (args.BreakOnError)
                    throw;
            }

            var json = JsonConvert.SerializeObject(updateLog);
            System.IO.File.WriteAllText(Path.Combine(args.Directory.FullName, "UpdateLog.txt"), json);

            logger.Info($"\n\nSynchronisation of fragments finished - {DateTime.Now}.\n");
            logger.Info($"Total documents processed {updateCounter.DocumentCount}, {updateCounter.DocumentUpdateCount} templates were updated\n");
            logger.Info($"Total fragments processed {updateCounter.ListItemCount}, {updateCounter.ListItemUpdateCount} items were updated\n\n");

        }

        private void ClearFolder(DirectoryInfo directory)
        {
            foreach (var file in directory.GetFiles())
            {
                file.Delete();
            }

            foreach (var dir in directory.GetDirectories())
            {
                dir.Delete(true);
            }
        }
    }
}
