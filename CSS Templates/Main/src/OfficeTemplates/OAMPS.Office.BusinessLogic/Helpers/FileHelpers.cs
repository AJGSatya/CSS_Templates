using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace OAMPS.Office.BusinessLogic.Helpers
{
    public class FileHelpers
    {
        public static bool IsLocalFile(string fileUrl)
        {
            return fileUrl.ToLower().StartsWith("file://");
        }

        public static string ExpandLocalFilePath(string fileUrl)
        {
            if (!IsLocalFile(fileUrl))
                return fileUrl;
            return Environment.ExpandEnvironmentVariables(fileUrl.Substring(7));
        }

        public static byte[] GetFileStreamWithMetadata(string path, out Dictionary<string, object> fieldValues, string fieldValuesFileEnding = "_fvs.txt")
        {
            fieldValues = new Dictionary<string, object>();

            var content = File.ReadAllBytes(path);

            var metaPath = path + fieldValuesFileEnding;
            var altMetaPath = BusinessLogic.Helpers.FileHelpers.ReplaceLastOccurrence(metaPath, "_jpg.jpg", ".jpg");
            altMetaPath = BusinessLogic.Helpers.FileHelpers.ReplaceLastOccurrence(altMetaPath, "_png.jpg", ".png");
            if (File.Exists(metaPath))
            {
                var metadata = File.ReadAllText(metaPath);
                fieldValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(metadata);
            }
            else if (File.Exists(altMetaPath))
            {
                var metadata = File.ReadAllText(altMetaPath);
                fieldValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(metadata);
            }
            return content;
        }

        public static T GetFileOfType<T>(string path) where T : class
        {
            if (File.Exists(path))
            {
                var data = File.ReadAllText(path);
                var obj = JsonConvert.DeserializeObject<T>(data);
                return obj;
            }
            else
            {
                return default(T);
            }
        }

        public static string GetThumbnailPath(string path)
        {
            if (path.EndsWith(".jpg", StringComparison.CurrentCultureIgnoreCase))
            {
                return ReplaceLastOccurrence(path, ".jpg", "_jpg.jpg");
            }
            else if (path.EndsWith(".png", StringComparison.CurrentCultureIgnoreCase))
            {
                return ReplaceLastOccurrence(path, ".png", "_png.jpg");
            }
            return path;
        }

        public static string ReplaceLastOccurrence(string source, string find, string replace)
        {
            var place = source.LastIndexOf(find);

            if (place == -1)
                return source;

            string result = source.Remove(place, find.Length).Insert(place, replace);
            return result;
        }
    }
}
