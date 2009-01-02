using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    static class TagExtension
    {
        private const string TagPrefix = "PowerPointLaTeX_";

        public static void PurgeAddInTags(this Tags tags, string prefix)
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

        public static IEnumerable<string> GetAddInNames(this Tags tags, string prefix)
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

        public static void SetAddInTag(this Tags tags, string name, string value)
        {
            tags.Add(TagPrefix + name, value);
        }

        public static string GetAddInTag(this Tags tags, string name)
        {
            return tags[TagPrefix + name];
        }

        public static void ClearAddInTag(this Tags tags, string name)
        {
            tags.Delete(TagPrefix + name);
        }

        public static LaTeXTags LaTeXTags(this Shape shape)
        {
            return new LaTeXTags(shape);
        }

        public static CacheTags CacheTags(this Presentation presentation)
        {
            return new CacheTags(presentation);
        }

        public static SettingsTags SettingsTags(this Presentation presentation)
        {
            return new SettingsTags(presentation);
        }
    }

    delegate void ValueChangedEventHandler<T>(object sender, T value);

    abstract class AddInTagBase<T>
    {
        public event ValueChangedEventHandler<T> ValueChanged;

        protected string name;
        protected Tags tags;

        public abstract T value
        {
            get;
            set;
        }

        protected string rawValue
        {
            get
            {
                return tags.GetAddInTag(name);
            }
            set
            {
                tags.SetAddInTag(name, value);

                FireEvent();
            }
        }

        private void FireEvent()
        {
            // make it thread-safe
            ValueChangedEventHandler<T> handler = ValueChanged;
            if (handler != null)
            {
                handler(this, this.value);
            }
        }

        public AddInTagBase(Tags tags, string name)
        {
            this.tags = tags;
            this.name = name;
        }

        public static implicit operator T(AddInTagBase<T> property)
        {
            return property.value;
        }

        public void Clear()
        {
            tags.ClearAddInTag(name);
            FireEvent();
        }
    }

    class AddInTagBool : AddInTagBase<bool>
    {
        public override bool value
        {
            get
            {
                return Helper.ParseBool(rawValue);
            }
            set
            {
                rawValue = value.ToString();
            }
        }

        public AddInTagBool(Tags tags, string name)
            : base(tags, name)
        {
        }
    }

    class AddInTagInt : AddInTagBase<int>
    {
        public override int value
        {
            get
            {
                return Helper.ParseInt(rawValue);
            }
            set
            {
                rawValue = value.ToString();
            }
        }

        public AddInTagInt(Tags tags, string name)
            : base(tags, name)
        {
        }
    }


    class AddInTagEnum<T> : AddInTagBase<T>
    {
        public override T value
        {
            get
            {
                T result = default(T);
                try
                {
                    result = (T) Enum.Parse(typeof(T), rawValue);
                }
                catch
                {
                }
                return result;
            }
            set
            {
                rawValue = value.ToString();
            }
        }

        public AddInTagEnum(Tags tags, string name)
            : base(tags, name)
        {
        }
    }

    class AddInTagString : AddInTagBase<string>
    {
        public override string value
        {
            get
            {
                return rawValue;
            }
            set
            {
                rawValue = value;
            }
        }

        public AddInTagString(Tags tags, string name)
            : base(tags, name)
        {
        }
    }

    class AddInTagByteArray : AddInTagBase<byte[]>
    {
        public override byte[] value
        {
            get
            {
                return Convert.FromBase64String(rawValue);
            }
            set
            {
                rawValue = Convert.ToBase64String(value);
            }
        }

        public AddInTagByteArray(Tags tags, string name)
            : base(tags, name)
        {
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
        internal static int ParseInt(string text)
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

        internal static bool ParseBool(string text)
        {
            bool value = false;
            try
            {
                value = bool.Parse(text);
            }
            catch
            {
            }
            return value;
        }
    }
}
