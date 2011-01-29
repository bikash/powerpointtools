using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLaTeX
{
    class LocalCache : ICacheStorage
    {
        private Dictionary<string, CacheEntry> cache = new Dictionary<string, CacheEntry>();
        private ICacheStorage masterStorage;

        public ICacheStorage MasterStorage
        {
            get { return masterStorage; }
        }

        public LocalCache() {
            this.masterStorage = null;
        }

        public LocalCache(ICacheStorage masterStorage)
        {
            this.masterStorage = masterStorage;
        }

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
            else if (masterStorage != null)
            {
                return masterStorage.Get(code);
            }
            else
            {
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
