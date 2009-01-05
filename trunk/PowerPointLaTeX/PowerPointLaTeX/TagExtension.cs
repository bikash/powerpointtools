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

        public T value
        {
            get {
                return FromString(tags.GetAddInTag(name));
            }
            set {
                if (!value.Equals(default(T)))
                {
                    tags.SetAddInTag(name, ToString(value));
                } else
                {
                    tags.ClearAddInTag(name);
                }

                FireEvent();
            }
        }

        protected abstract T FromString(string rawValue);
        protected virtual string ToString(T value) {
            return value.ToString();
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
        public AddInTagBool(Tags tags, string name)
            : base(tags, name)
        {
        }

        protected override bool FromString(string rawValue)
        {
            return Helper.ParseBool(rawValue);
        }
    }

    class AddInTagInt : AddInTagBase<int>
    {
        public AddInTagInt(Tags tags, string name)
            : base(tags, name)
        {
        }

        protected override int FromString(string rawValue)
        {
            return Helper.ParseInt(rawValue);
        }
    }


    class AddInTagEnum<T> : AddInTagBase<T>
    {
        public AddInTagEnum(Tags tags, string name)
            : base(tags, name)
        {
        }

        protected override T FromString(string rawValue)
        {
            if( Enum.IsDefined(typeof(T), rawValue)) {
                return (T) Enum.Parse(typeof(T), rawValue);
            }
            return default(T);
        }
    }

    class AddInTagString : AddInTagBase<string>
    {
        public AddInTagString(Tags tags, string name)
            : base(tags, name)
        {
        }

        protected override string FromString(string rawValue)
        {
            return rawValue;
        }
    }

    class AddInTagByteArray : AddInTagBase<byte[]>
    {
        public AddInTagByteArray(Tags tags, string name)
            : base(tags, name)
        {
        }

        protected override byte[] FromString(string rawValue)
        {
            return Convert.FromBase64String(rawValue);
        }

        protected override string ToString(byte[] value)
        {
            return Convert.ToBase64String(value);
        }
    }

    // TODO: move this somewhere else, too [12/31/2008 Andreas]
    static class Helper
    {
        internal static int ParseInt(string text)
        {
            int value;
            if( int.TryParse(text, out value)) {
                return value;
            }
            return 0;
        }

        internal static bool ParseBool(string text)
        {
            bool value;
            if( bool.TryParse(text, out value)) {
                return value;
            }
            return false;
        }
    }
}
