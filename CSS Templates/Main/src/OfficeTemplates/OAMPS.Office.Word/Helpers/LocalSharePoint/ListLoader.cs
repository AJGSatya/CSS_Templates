using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Loggers;
using OAMPS.Office.BusinessLogic.Models;
using OAMPS.Office.Word.Models;
using OAMPS.Office.Word.Models.Local;
using OAMPS.Office.Word.Properties;
using Type = OAMPS.Office.BusinessLogic.Interfaces.Logging.Type;

namespace OAMPS.Office.Word.Helpers.LocalSharePoint
{
    public class ListLoader
    {
        private EventViewerLogger _logger = new EventViewerLogger();

        private UpdateLog _listLookups;
        private readonly string _basePath = Environment.ExpandEnvironmentVariables(Settings.Default.LocalFullPathRoot);

        public ListLoader()
        {
            _listLookups = LocalCache.Get<UpdateLog>("LocalSharePoint.List.UpdateLog", () =>
            {
                var path = Path.Combine(_basePath, "UpdateLog.txt");
                var data = File.ReadAllText(path);
                var log = JsonConvert.DeserializeObject<UpdateLog>(data);
                return log;
            });
        }

        public string GetLocalPath(string listTitle)
        {
            _logger.Log("Loading list from local cache: " + listTitle, Type.Information);

            var listLookup = _listLookups.Updates.FirstOrDefault(x => x.Source.Equals(listTitle, StringComparison.InvariantCultureIgnoreCase));
            if (listLookup == null)
            {
                _logger.Log("List not found in local cache: " + listTitle, Type.Error);
                return null;
            }
                
            return Path.Combine(_basePath, listLookup.Destination);
        }

        public ListItems GetItems(string listTitle)
        {
            _logger.Log("Gettings items from local cache: " + listTitle, Type.Information);

            return LocalCache.Get<ListItems>("LocalSharePoint.List.GetItems." + listTitle.Replace(" ", ""), () => LoadListItems(listTitle));
        }

        public Dictionary<string, object> GetListSettings(string listTitle)
        {
            var settingsFile = Path.Combine(GetLocalPath(listTitle), "list-settings.txt");
            return FileHelpers.GetFileOfType<Dictionary<string, object>>(settingsFile);
        }

        private ListItems LoadListItems(string listTitle)
        {
            var listLookup = _listLookups.Updates.FirstOrDefault(x => x.Source.Equals(listTitle, StringComparison.InvariantCultureIgnoreCase));
            if (listLookup == null)
                return default(ListItems);

            var listFolder = GetLocalPath(listTitle);
            var listFile = Path.Combine(listFolder, "_AllListContent" + Settings.Default.FieldValuesFileEnding);
            var listVersionFile = Path.Combine(listFolder, "_AllListContentVersionInfo.txt");

            var versionInfo = FileHelpers.GetFileOfType<ListVersion>(listVersionFile);
            if (versionInfo == null || versionInfo.Version < listLookup.Version || !File.Exists(listFile))
            {
                _logger.Log("List items changed, recompiling items: " + listTitle, Type.Information);

                var items = CombineListItems(listFolder, listFile);

                versionInfo = new ListVersion()
                {
                    LastModified = listLookup.LastModified,
                    Title = listLookup.Source,
                    Version = listLookup.Version
                };
                var output = JsonConvert.SerializeObject(versionInfo);
                File.WriteAllText(listVersionFile, output);

                return items;
            }

            return FileHelpers.GetFileOfType<ListItems>(listFile);
        }

        private ListItems CombineListItems(string listFolder, string listFile)
        {
            var listItems = new ListItems();
            var dirInfo = new DirectoryInfo(listFolder);
            foreach (var file in dirInfo.GetFiles("*" + Settings.Default.FieldValuesFileEnding))
            {
                if (file.FullName.Equals(listFile, StringComparison.InvariantCultureIgnoreCase))
                    continue;

                var item = ListItem.Parse(file.FullName);
                listItems.Add(item);
            }

            var output = JsonConvert.SerializeObject(listItems);
            File.WriteAllText(listFile, output);

            return listItems;
        }
    }
}
