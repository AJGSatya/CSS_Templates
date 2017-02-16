using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Threading.Tasks;

namespace OAMPS.Office.Word.Models.Local
{
    public class LocalCache
    {
        private static object _lock = new object();
        private static ObjectCache cache = MemoryCache.Default;

        private const string KeyPrefix = "OAMPS.LOCAL:";

        public static T Get<T>(string key, Func<T> retrieve = null) where T : class
        {
            var internalKey = GetInternalKey(key);
            if (cache.Contains(internalKey))
            {
                return cache[internalKey] as T;
            }

            T res;
            lock (_lock)
            {
                if (cache.Contains(internalKey))
                {
                    return cache[internalKey] as T;
                }

                if (retrieve != null)
                {
                    res = retrieve();
                    if (res != null)
                        cache[internalKey] = res;
                }
                else
                {
                    return default(T);
                }
            }
            return res;

        }

        private static string GetInternalKey(string key) => $"{KeyPrefix}{key}";
    }
}
