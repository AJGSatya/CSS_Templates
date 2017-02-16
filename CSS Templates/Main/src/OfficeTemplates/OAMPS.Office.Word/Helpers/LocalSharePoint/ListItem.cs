using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace OAMPS.Office.Word.Helpers.LocalSharePoint
{
    public class ListItem : Dictionary<string, object>
    {
        public static ListItem Parse(string path)
        {
            var data = File.ReadAllText(path);
            var obj = JsonConvert.DeserializeObject<ListItem>(data);
            return obj;
        }
    }
}