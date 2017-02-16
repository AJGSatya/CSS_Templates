using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSTemplates.TemplateSyncTool.Settings;
using log4net;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OAMPS.Office.BusinessLogic.Models;

namespace CSSTemplates.TemplateSyncTool.Synchronise
{
    abstract class ItemSynchroniserBase
    {
        public void Synchronise(SharePointClientContext spContext, Fragment fragment, UpdateCounter updateCounter, UpdateLog updateLog, log4net.ILog logger, Arguments args)
        {
            var list = spContext.Web.Lists.GetByTitle(fragment.Source);
            var items = list.GetItems(new CamlQuery());
            LoadListItems(spContext, items);
            spContext.ExecuteQuery();

            var noUpdateRequiredCount = 0;
            var updatedCount = 0;

            var destinationPath = args.Directory.FullName;
            if (fragment.Destination.Length > 0)
            {
                var destination = args.Directory.CreateSubdirectory(fragment.Destination);
                destinationPath = destination.FullName;
            }

            logger.Info($"\n\nProcessing {fragment.Source} list settings...");
            GetListSettings(spContext, list, fragment.Source, destinationPath);

            logger.Info($"\n\nProcessing {fragment.Source} into {destinationPath}, found {items.Count} fragments");

            foreach (var item in items)
            {
                if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                    continue;

                HandleListItem(spContext, logger, args, destinationPath, item, ref updatedCount, ref noUpdateRequiredCount);

                var fileName = GetFileNameForFieldValue(item, destinationPath);
                SaveItemFieldValues(logger, args.Overwrite, item, fileName, ref updatedCount, ref noUpdateRequiredCount);
            }

            logger.Info($"\nFinished processing {fragment.Source}, {noUpdateRequiredCount} items were up to date, {updatedCount} items were updated");
            UpdateCounters(updateCounter, noUpdateRequiredCount, updatedCount);

            var log = updateLog.Updates.FirstOrDefault(x => x.Source.Equals(fragment.Source, StringComparison.InvariantCultureIgnoreCase));
            if (log == null)
            {
                log = new UpdateLog.UpdateLogEntry()
                {
                    Source = fragment.Source,
                    Destination = fragment.Destination,
                    LastModified = DateTime.Now,
                    Version = 1
                };
                updateLog.Updates.Add(log);
            }
            else
            {
                log.Destination = fragment.Destination;
                if (updatedCount > 0)
                {
                    log.LastModified = DateTime.Now;
                    log.Version++;
                }
            }
        }

        public virtual void GetListSettings(SharePointClientContext spContext, List list, string title, string destinationPath)
        {
            // extend this when needed to be more generic...

            if (title != "OAMPS Branding")
                return;

            var values = new Dictionary<string, object>();

            var choices = spContext.CastTo<FieldChoice>(list.Fields.GetByTitle("Speciality"));
            spContext.Load(choices, f => f.Choices, f => f.DefaultValue);
            spContext.ExecuteQuery();
            var specialities = choices.Choices;

            values.Add("Speciality.DefaultValue", choices.DefaultValue);
            values.Add("Speciality.Choices", specialities);
            var json = JsonConvert.SerializeObject(values);
            System.IO.File.WriteAllText(Path.Combine(destinationPath, "list-settings.txt"), json);
        }

        public virtual string GetFileNameForFieldValue(ListItem item, string destinationPath)
        {
            var title = item.FieldValues["Title"] as string;
            var fileName = Path.Combine(destinationPath, Sanitize.GetValidFileName(title) + ConfigurationManager.AppSettings["FieldValuesFileEnding"]);
            return fileName;
        }

        public virtual void HandleListItem(SharePointClientContext spContext, ILog logger, Arguments args, string destinationPath, ListItem item, ref int updatedCount, ref int noUpdateRequiredCount)
        {
            throw new NotImplementedException();
        }

        public virtual void LoadListItems(SharePointClientContext spContext, ListItemCollection items)
        {
            spContext.Load(items,
                t => t.Include(doc => doc.ParentList.RootFolder.ServerRelativeUrl),
                t => t.Include(doc => doc.ParentList.ParentWebUrl),
                t => t);
        }

        private void SaveItemFieldValues(ILog logger, bool overwrite, ListItem item, string fileName, ref int updatedCount, ref int noUpdateRequiredCount)
        {
            if (overwrite || (DateTime)item["Modified"] > System.IO.File.GetLastWriteTime(fileName))
            {
                logger.Info($"Updating: {fileName}");
                var values = item.FieldValues;

                // persist additional context - i.e. web url
                if (item.ParentList.ServerObjectIsNull == false)
                {
                    values.Add("ParentList.RootFolder.ServerRelativeUrl", item.ParentList.RootFolder.ServerRelativeUrl);
                    values.Add("ParentList.ParentWebUrl", item.ParentList.ParentWebUrl);
                }

                var json = JsonConvert.SerializeObject(values);
                System.IO.File.WriteAllText(fileName, json);
                updatedCount++;
            }
            else
            {
                logger.Info($"Skipping: {fileName}");
                noUpdateRequiredCount++;
            }
        }

        public virtual void UpdateCounters(UpdateCounter updateCounter, int noUpdateRequiredCount, int updatedCount)
        {
            updateCounter.ListItemCount += (noUpdateRequiredCount + updatedCount);
            updateCounter.ListItemUpdateCount += updatedCount;
        }
    }

}
