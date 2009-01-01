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
            private string code;
            private Presentation presentation;

            public CacheEntry(Presentation presentation, string code)
            {
                this.presentation = presentation;
                // because the tags system converts names to uppercase, we have to use a different format
                // -> convert it to hex
                StringBuilder hexData = new StringBuilder();
                foreach (byte c in code)
                {
                    hexData.Append( c.ToString("X2") );
                }
                this.code = hexData.ToString();
            }

            public int RefCounter
            {
                get
                {
                    return Helper.ParseIntToString(presentation.GetTag("RefCounter#" + code));
                }
                private set
                {
                    presentation.SetTag("RefCounter#" + code, value.ToString());
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
                    return Convert.FromBase64String(presentation.GetTag("CacheContent#" + code));
                }
                private set
                {
                    presentation.SetTag("CacheContent#" + code, Convert.ToBase64String(value));
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

            public void Clear()
            {
                Debug.Assert(RefCounter <= 1);
                presentation.ClearTag("RefCounter#" + code);
                presentation.ClearTag("CacheContent#" + code);
            }
        }

        private Presentation presentation;

        public CacheTags(Presentation presentation)
        {
            this.presentation = presentation;
        }

        public void PurgeAll()
        {
            presentation.PurgeTags("RefCounter#");
            presentation.PurgeTags("CacheContent#");
        }

        public void PurgeUnused()
        {
            string refCounterPrefix = "RefCounter#";
            IEnumerable<string> names = presentation.GetTagNames(refCounterPrefix);
            foreach (string name in names)
            {
                CacheEntry entry = new CacheEntry(presentation, name.Substring(refCounterPrefix.Length));
                if (entry.RefCounter <= 1)
                {
                    entry.Clear();
                }
            }
        }

        public CacheEntry this[string code]
        {
            get { return new CacheEntry(presentation, code); }
        }
    }

}
