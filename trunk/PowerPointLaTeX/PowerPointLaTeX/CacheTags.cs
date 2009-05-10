#region Copyright Notice
// This file is part of PowerPoint LaTeX.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint LaTeX is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint LaTeX is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

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
            // refCounter is 1 + #references
            // 0 means that has never been accessed and 1 means that the entry is cached but currently unused
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

            public void ReleaseIfUsed()
            {
                if( RefCounter == 0 ) {
                    // unused
                    return;
                }
                // otherwise the refcounter must be at least 2
                Debug.Assert(RefCounter >= 2);
                if (--RefCounter < 1)
                {
                    // even if something goes wrong, keep it as 'cached'
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
