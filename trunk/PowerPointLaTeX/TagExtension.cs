using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    // exceptions?
    // TODO: change this to be object-centric (wrap the specific fields into objects) [12/31/2008 Andreas]
    static class TagExtension
    {
        private const string TagPrefix = "PowerPointLaTeX_";

        public static void SetTag(this Shape shape, string name, string value)
        {
            shape.Tags.Add(TagPrefix + name, value);
        }

        public static string GetTag(this Shape shape, string name)
        {
            return shape.Tags[TagPrefix + name];
        }

        public static void ClearTag(this Shape shape, string name)
        {
            shape.Tags.Delete(TagPrefix + name);
        }

        public static LaTeXTags LaTeXTags(this Shape shape)
        {
            return new LaTeXTags(shape);
        }
    }

    /*
        class TagProperty<T> where T : struct {
            private Shape shape;
            private string name;

            public static implicit operator T
        }*/

    static class Helper
    {
        internal static int ParseIntToString(string text)
        {
            int value = 0;
            try
            {
                value = int.Parse(text);
            }
            catch
            {
            }
            return value;
        }
    }

    class LaTeXEntry
    {
        private Shape shape;
        private int index;

        public LaTeXEntry(Shape shape, int index)
        {
            this.shape = shape;
            this.index = index;
        }

        public void Clear() {
            shape.ClearTag("Entry[" + index + "].Code");
            shape.ClearTag("Entry[" + index + "].StartIndex");
            shape.ClearTag("Entry[" + index + "].Length");
            shape.ClearTag("Entry[" + index + "].ShapeID");
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

        public int ShapeID
        {
            get
            {
                return Helper.ParseIntToString(shape.GetTag("Entry[" + index + "].ShapeID"));
            }
            set
            {
                shape.SetTag("Entry[" + index + "].ShapeID", value.ToString());
            }
        }
    }

    class LaTeXEntries
    {
        private Shape shape;

        public LaTeXEntries(Shape shape)
        {
            this.shape = shape;
        }

        public void Clear() {
            for( int i = 0 ; i < Length ; i++ ) {
                this[i].Clear();
            }
            shape.ClearTag("Entry.Length");
        }

        public LaTeXEntry this[int index]
        {
            get
            {
                if( index >= Length ) {
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
    }

    class LaTeXTags
    {
        private Shape shape;

        public LaTeXTags(Shape shape)
        {
            this.shape = shape;
        }

        public void Clear() {
            shape.ClearTag("Code");
            shape.ClearTag("Type");
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
                LaTeXTool.EquationType type = LaTeXTool.EquationType.Inline;
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

        public LaTeXEntries Entries
        {
            get
            {
                return new LaTeXEntries(shape);
            }
        }
    }
}
