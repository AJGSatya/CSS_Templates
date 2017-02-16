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
using OAMPS.Office.BusinessLogic.Helpers;

namespace CSSTemplates.TemplateSyncTool.Synchronise
{
    class FileSynchroniser : ItemSynchroniserBase
    {
        public override void HandleListItem(SharePointClientContext spContext, ILog logger, Arguments args, string destinationPath, ListItem item, ref int updatedCount, ref int noUpdateRequiredCount)
        {
            var fileName = Path.Combine(destinationPath, item.File.Name);
            if (args.Overwrite || item.File.TimeLastModified > System.IO.File.GetLastWriteTime(fileName))
            {
                logger.Info($"Updating: {fileName}");
                using (var fileStream = System.IO.File.Create(fileName))
                {
                    var spFileStream = item.File.OpenBinaryStream();
                    spContext.ExecuteQuery();

                    spFileStream.Value.CopyTo(fileStream);
                    updatedCount++;
                }

                if (fileName.EndsWith(".jpg", StringComparison.InvariantCultureIgnoreCase) || fileName.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase))
                {
                    var thumbName = FileHelpers.GetThumbnailPath(fileName);
                    using (var thumbStream = System.IO.File.Create(thumbName))
                    {
                        var fileUrl = item.File.ServerRelativeUrl;
                        var thumbUrl = FileHelpers.ReplaceLastOccurrence(FileHelpers.GetThumbnailPath(fileUrl), "/", "/_t/");
                        var thumb = spContext.Web.GetFileByServerRelativeUrl(thumbUrl);
                        spContext.Load(thumb);
                        var spThumbStream = thumb.OpenBinaryStream();
                        spContext.ExecuteQuery();
                        spThumbStream.Value.CopyTo(thumbStream);
                    }
                }
            }
            else
            {
                noUpdateRequiredCount++;
                logger.Info($"Skipping: {fileName}");
            }
        }

        public override string GetFileNameForFieldValue(ListItem item, string destinationPath)
        {
            var fileName = Path.Combine(destinationPath, item.File.Name);
            return fileName + ConfigurationManager.AppSettings["FieldValuesFileEnding"];
        }

        public override void LoadListItems(SharePointClientContext spContext, ListItemCollection items)
        {
            spContext.Load(items, t => t.Include(doc => doc.File),
                t => t.Include(doc => doc.ParentList.RootFolder.ServerRelativeUrl),
                t => t.Include(doc => doc.ParentList.ParentWebUrl),
                t => t);
        }

        public override void UpdateCounters(UpdateCounter updateCounter, int noUpdateRequiredCount, int updatedCount)
        {
            updateCounter.DocumentCount += (noUpdateRequiredCount + updatedCount);
            updateCounter.DocumentUpdateCount += updatedCount;
        }
    }

}
