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

        private static void InternalPurgeTags(Tags tags, string prefix)
        {
            int i = 1;
            string namePrefix = TagPrefix + prefix;
            while (i <= tags.Count)
            {
                string name = tags.Name(i);
                if (name.StartsWith(namePrefix, StringComparison.OrdinalIgnoreCase))
                {
                    tags.Delete(name);
                }
                else
                {
                    i += 1;
                }
            }
        }

        private static IEnumerable<string> InternalGetTagNames(Tags tags, string prefix)
        {
            List<string> names = new List<string>();
            string namePrefix = TagPrefix + prefix;
            for (int i = 1; i <= tags.Count; i++)
            {
                if (tags.Name(i).StartsWith(namePrefix, StringComparison.OrdinalIgnoreCase))
                {
                    string name = tags.Name(i).Substring(namePrefix.Length);
                    names.Add(name);
                }
            }

            return names;
        }

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

        public static void PurgeTags(this Shape shape, string prefix)
        {
            InternalPurgeTags(shape.Tags, prefix);
        }

        public static IEnumerable<string> GetTagNames(this Shape shape, string prefix)
        {
            return InternalGetTagNames(shape.Tags, prefix);
        }

        // TODO: wtf? how can I get rid of this duplicate code? [12/31/2008 Andreas]
        public static void SetTag(this Presentation presentation, string name, string value)
        {
            presentation.Tags.Add(TagPrefix + name, value);
        }

        public static string GetTag(this Presentation presentation, string name)
        {
            return presentation.Tags[TagPrefix + name];
        }

        public static void ClearTag(this Presentation presentation, string name)
        {
            presentation.Tags.Delete(TagPrefix + name);
        }

        public static void PurgeTags(this Presentation presentation, string prefix)
        {
            InternalPurgeTags(presentation.Tags, prefix);
        }

        public static IEnumerable<string> GetTagNames(this Presentation presentation, string prefix)
        {
            return InternalGetTagNames(presentation.Tags, prefix);
        }

        public static LaTeXTags LaTeXTags(this Shape shape)
        {
            return new LaTeXTags(shape);
        }

        public static CacheTags CacheTags(this Presentation presentation)
        {
            return new CacheTags(presentation);
        }
    }

    /*
        class TagProperty<T> where T : struct {
            private Shape shape;
            private string namePrefix;

            public static implicit operator T
        }*/

    // TODO: move this somewhere else, too [12/31/2008 Andreas]
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
}
