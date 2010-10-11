using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLaTeX
{
    class LocalCache : ICacheStorage
    {
        Dictionary<string, CacheEntry> cache = new Dictionary<string, CacheEntry>();

        #region ICacheStorage Members

        public void PurgeAll()
        {
            cache.Clear();
        }

        public CacheEntry? Get(string code)
        {
            CacheEntry entry = new CacheEntry();
            if( cache.TryGetValue(code, out entry) ) {
                return entry;
            }
            else {
                return null;
            }
        }

        public void Set(string code, CacheEntry? entry)
        {
            if( entry.HasValue ) {
                cache.Remove(code);
            }
            cache.Add(code, entry.Value);
        }

        #endregion
    }
}
