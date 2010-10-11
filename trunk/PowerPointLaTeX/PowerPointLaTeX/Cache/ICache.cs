using System;
namespace PowerPointLaTeX
{
    struct CacheEntry
    {
        public int BaselineOffset;
        public byte[] Content;
        public float PixelsPerEmHeight;
    }

    interface ICacheStorage
    {
        void PurgeAll();
        CacheEntry? Get(string code);
        void Set(string code, CacheEntry? entry);
    }

    static class Cache
    {
        public static void Store(ICacheStorage storage, string code, CacheEntry entry) {
            CacheEntry? currentEntry = storage.Get(code);
            // don't overwrite if the new content is worse than the old one
            // convert to int to avoid floating point issues [5/24/2010 Andreas]
            if (currentEntry != null && (int) currentEntry.Value.PixelsPerEmHeight >= (int) entry.PixelsPerEmHeight)
            {
                return;
            }

            storage.Set(code, entry);
        }

        public static CacheEntry? Query(ICacheStorage storage, string code) {
            return storage.Get(code);
        }
    }
}
