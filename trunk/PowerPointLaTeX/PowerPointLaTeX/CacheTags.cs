using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PowerPointLaTeX
{
    class CacheTags
    {
        public class CacheEntry
        {
            private AddInTagInt refCounter;
            private AddInTagByteArray content;

            public CacheEntry(Tags tags, string code)
            {
                refCounter = new AddInTagInt(tags, "RefCounter#" + code);
                content = new AddInTagByteArray(tags, "CacheContent#" + code);
            }

            public void Clear()
            {
                Debug.Assert(RefCounter <= 1);
                refCounter.Clear();
                content.Clear();
            }
            public int RefCounter
            {
                get
                {
                    return refCounter;
                }
                private set
                {
                    refCounter.value = value;
                }
            }

            public bool IsCached()
            {
                return RefCounter > 0;
            }

            public byte[] Content
            {
                get
                {
                    return content;
                }
                private set
                {
                    content.value = value;
                }
            }

            public void Release()
            {
                Debug.Assert(RefCounter > 1);
                if (--RefCounter < 1)
                {
                    RefCounter = 1;
                }
            }

            public void Store(byte[] data)
            {
                Debug.Assert(!IsCached());
                Content = data;
                RefCounter = 2;
            }

            public byte[] Use()
            {
                RefCounter++;
                return Content;
            }

        }

        private Tags tags;

        public CacheTags(Presentation presentation)
        {
            tags = presentation.Tags;
        }

        public void PurgeAll()
        {
            tags.PurgeAddInTags("RefCounter#");
            tags.PurgeAddInTags("CacheContent#");
        }

        public void PurgeUnused()
        {
            string refCounterPrefix = "RefCounter#";
            IEnumerable<string> names = tags.GetAddInNames(refCounterPrefix);
            foreach (string name in names)
            {
                CacheEntry entry = new CacheEntry(tags, name);
                if (entry.RefCounter <= 1)
                {
                    entry.Clear();
                }
            }
        }

        public CacheEntry this[string code]
        {
            get {
                // because the tags system converts names to uppercase, we have to use a different format
                // -> convert it to hex
                StringBuilder hexData = new StringBuilder();
                foreach (byte c in code)
                {
                    hexData.Append(c.ToString("X2"));
                }
                string encodedCode = hexData.ToString();

                return new CacheEntry(tags, encodedCode);
            }
        }
    }

}
