using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace CustomExtensions
{
    // exceptions?
    // TODO: change this to be object-centric (wrap the specific fields into objects) [12/31/2008 Andreas]
    static class TagExtension
    {
        private const string TagPrefix = "PowerPointLaTeX_";

        public static void AddElementTag(this Shape shape, string name, string value)
        {
            shape.Tags.Add(TagPrefix + name, value);
        }

        public static string GetElementTag(this Shape shape, string name)
        {
            return shape.Tags[name];
        }

        public static void AddArrayTag(this Shape shape, string name, int index, string value)
        {
            shape.Tags.Add(TagPrefix + name + "[" + value + "]", value);
        }

        public static string GetArrayTag(this Shape shape, string name, int index)
        {
            return shape.Tags[name + "[" + index + "]"];
        }
    }
}
