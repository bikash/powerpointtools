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

namespace PowerPointLaTeX
{
    class LaTeXEntry
    {
        public AddInTagString Code;
        public AddInTagInt StartIndex;
        public AddInTagInt Length;
        public AddInTagInt ShapeId;

        public LaTeXEntry(Tags tags, int index)
        {
            this.Code = new AddInTagString(tags, "Entry[" + index + "].Code");
            this.StartIndex = new AddInTagInt(tags, "Entry[" + index + "].StartIndex");
            this.Length = new AddInTagInt(tags, "Entry[" + index + "].Length");
            this.ShapeId = new AddInTagInt(tags, "Entry[" + index + "].ShapeId");
        }

        public void Clear()
        {
            Code.Clear();
            StartIndex.Clear();
            Length.Clear();
            ShapeId.Clear();
        }
    }

    class LaTeXEntries : IEnumerable<LaTeXEntry>
    {
        private Tags tags;
        public AddInTagInt Length;

        public LaTeXEntries(Tags tags)
        {
            this.tags = tags;
            this.Length = new AddInTagInt(tags, "Entry.Length");
        }

        public void Clear()
        {
            for (int i = 0; i < Length; i++)
            {
                this[i].Clear();
            }
            Length.Clear();
        }

        public LaTeXEntry this[int index]
        {
            get
            {
                if (index >= Length)
                {
                    Length.value = index + 1;
                }
                return new LaTeXEntry(tags, index);
            }
        }

        #region IEnumerable<LaTeXEntries> Members
        private class Enumerator : IEnumerator<LaTeXEntry>
        {
            private LaTeXEntries parent;
            private int index = -1;

            public Enumerator(LaTeXEntries parent)
            {
                this.parent = parent;
            }

            #region IEnumerator<LaTeXEntry> Members

            public LaTeXEntry Current
            {
                get
                {
                    if (index < 0 || index >= parent.Length)
                    {
                        throw new InvalidOperationException();
                    }
                    return parent[index];
                }
            }

            #endregion

            #region IDisposable Members

            public void Dispose()
            {
            }

            #endregion

            #region IEnumerator Members

            object System.Collections.IEnumerator.Current
            {
                get { return Current as object; }
            }

            public bool MoveNext()
            {
                index++;
                return index < parent.Length;
            }

            public void Reset()
            {
                index = 0;
            }

            #endregion
        }

        public IEnumerator<LaTeXEntry> GetEnumerator()
        {
            return new Enumerator(this);
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new Enumerator(this);
        }

        #endregion
    }

    class LaTeXTags
    {
        public readonly AddInTagString Code;
        public readonly AddInTagEnum<EquationType> Type;
        public readonly AddInTagInt LinkID;
        
        public readonly LaTeXEntries Entries;

        public readonly AddInTagFloat OriginalWidth, OriginalHeight;

        public LaTeXTags(Shape shape)
        {
            Tags tags = shape.Tags;
            Code = new AddInTagString(tags, "Code");
            Type = new AddInTagEnum<EquationType>(tags, "Type");
            LinkID = new AddInTagInt(tags, "LinkID");
            Entries = new LaTeXEntries(tags);

            OriginalWidth = new AddInTagFloat(tags, "OriginalWidth");
            OriginalHeight = new AddInTagFloat(tags, "OriginalHeight");
        }

        public void Clear()
        {
            Code.Clear();
            Type.Clear();
            LinkID.Clear();
            Entries.Clear();
            OriginalWidth.Clear();
            OriginalHeight.Clear();
        }
    }
}
