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
    class CacheTags : ICacheStorage
    {
        public class CacheTagsEntry
        {
            private AddInTagBool cached;
            private AddInTagByteArray content;
            private AddInTagInt baselineOffset;
            private AddInTagFloat pixelsPerEmHeight;

            public CacheTagsEntry(Tags tags, string code)
            {
                cached = new AddInTagBool(tags, "Cached#" + code);
                content = new AddInTagByteArray(tags, "CacheContent#" + code);
                baselineOffset = new AddInTagInt( tags, "BaseLineOffset#" + code );
                pixelsPerEmHeight = new AddInTagFloat( tags, "PixelsPerEmHeight#" + code );
            }

            public void Clear()
            {
                cached.Clear();
                content.Clear();
                baselineOffset.Clear();
                pixelsPerEmHeight.Clear();
            }

            public bool Cached
            {
                get
                {
                    return cached;
                }
                set
                {
                    cached.value = value;
                }
            }

            public byte[] Content
            {
                get
                {
                    return content;
                }
                set
                {
                    content.value = value;
                }
            }

            public int BaselineOffset {
                get {
                    return baselineOffset;
                }
                set {
                    baselineOffset.value = value;
                }
            }

            public float PixelsPerEmHeight {
                get {
                    return pixelsPerEmHeight;
                }
                set {
                    pixelsPerEmHeight.value = value;
                }
            }
        }

        private Tags tags;

        public CacheTags(Presentation presentation)
        {
            tags = presentation.Tags;
        }

        #region ICacheStorage Members

        public void PurgeAll()
        {
            tags.PurgeAddInTags( "Cached#" );
            tags.PurgeAddInTags( "CacheContent#" );
            tags.PurgeAddInTags( "BaseLineOffset#" );
            tags.PurgeAddInTags( "PixelsPerEmHeight#" );
        }

        // TODO: public void PurgeUnused()
        private string encodeText(string text) {
            // because the tags system converts names to uppercase, we have to use a different format
            // -> convert it to hex
            StringBuilder hexData = new StringBuilder();
            foreach (byte c in text)
            {
                hexData.Append(c.ToString("X2"));
            }
            string encodedCode = hexData.ToString();
            return encodedCode;
        }

        private CacheTagsEntry getCacheTagsEntry( string code ) {
            return new CacheTagsEntry(tags, encodeText(code));
        }

        public CacheEntry? Get(string code)
        {
            CacheTagsEntry tagsEntry = getCacheTagsEntry(code);
            if( !tagsEntry.Cached ) {
                return null;
            }

            CacheEntry entry = new CacheEntry
                {
                    Content = tagsEntry.Content,
                    BaselineOffset = tagsEntry.BaselineOffset,
                    PixelsPerEmHeight = tagsEntry.PixelsPerEmHeight
                };
            return entry;
        }

        public void Set(string code, CacheEntry? entry)
        {
            CacheTagsEntry tagsEntry = getCacheTagsEntry(code);
            if( entry.HasValue ) {
                tagsEntry.Clear();
            }

            tagsEntry.Cached = true;
            tagsEntry.Content = entry.Value.Content;
            tagsEntry.BaselineOffset = entry.Value.BaselineOffset;
            tagsEntry.PixelsPerEmHeight = entry.Value.PixelsPerEmHeight;
        }

        #endregion
    }

}
