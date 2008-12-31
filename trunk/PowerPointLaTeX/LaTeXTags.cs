using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    class LaTeXEntry
    {
        private Shape shape;
        private int index;

        public LaTeXEntry(Shape shape, int index)
        {
            this.shape = shape;
            this.index = index;
        }

        public void Clear()
        {
            shape.ClearTag("Entry[" + index + "].Code");
            shape.ClearTag("Entry[" + index + "].StartIndex");
            shape.ClearTag("Entry[" + index + "].Length");
            shape.ClearTag("Entry[" + index + "].ShapeId");
        }

        public string Code
        {
            get
            {
                return shape.GetTag("Entry[" + index + "].Code");
            }
            set
            {
                shape.SetTag("Entry[" + index + "].Code", value);
            }
        }

        public int StartIndex
        {
            get
            {
                return Helper.ParseIntToString(shape.GetTag("Entry[" + index + "].StartIndex"));
            }
            set
            {
                shape.SetTag("Entry[" + index + "].StartIndex", value.ToString());
            }
        }

        public int Length
        {
            get
            {
                return Helper.ParseIntToString(shape.GetTag("Entry[" + index + "].Length"));
            }
            set
            {
                shape.SetTag("Entry[" + index + "].Length", value.ToString());
            }
        }

        public int ShapeId
        {
            get
            {
                return Helper.ParseIntToString(shape.GetTag("Entry[" + index + "].ShapeId"));
            }
            set
            {
                shape.SetTag("Entry[" + index + "].ShapeId", value.ToString());
            }
        }
    }

    class LaTeXEntries : IEnumerable<LaTeXEntry>
    {
        private Shape shape;

        public LaTeXEntries(Shape shape)
        {
            this.shape = shape;
        }

        public void Clear()
        {
            for (int i = 0; i < Length; i++)
            {
                this[i].Clear();
            }
            shape.ClearTag("Entry.Length");
        }

        public LaTeXEntry this[int index]
        {
            get
            {
                if (index >= Length)
                {
                    Length = index + 1;
                }
                return new LaTeXEntry(shape, index);
            }
        }

        public int Length
        {
            get
            {
                return Helper.ParseIntToString(shape.GetTag("Entry.Length"));
            }
            set
            {
                shape.SetTag("Entry.Length", value.ToString());
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
        private Shape shape;

        public LaTeXTags(Shape shape)
        {
            this.shape = shape;
        }

        public void Clear()
        {
            shape.ClearTag("Code");
            shape.ClearTag("Type");
            shape.ClearTag("ParentId");
            Entries.Clear();
        }

        public string Code
        {
            get
            {
                return shape.GetTag("Code");
            }
            set
            {
                shape.SetTag("Code", value);
            }
        }

        public LaTeXTool.EquationType Type
        {
            get
            {
                LaTeXTool.EquationType type = LaTeXTool.EquationType.None;
                try
                {
                    type = (LaTeXTool.EquationType) Enum.Parse(typeof(LaTeXTool.EquationType), shape.GetTag("Type"));
                }
                catch
                {
                }
                return type;
            }
            set
            {
                shape.SetTag("Type", value.ToString());
            }
        }

        public int ParentId
        {
            get
            {
                return Helper.ParseIntToString(shape.GetTag("ParentId"));
            }
            set
            {
                shape.SetTag("ParentId", value.ToString());
            }
        }

        public LaTeXEntries Entries
        {
            get
            {
                return new LaTeXEntries(shape);
            }
        }
    }
}
