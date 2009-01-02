using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PowerPointLaTeX
{
    class ReflectionPropertyDescriptor : PropertyDescriptor
    {
        private Object instance;
        private PropertyInfo propertyInfo;
        private object[] index;

        public ReflectionPropertyDescriptor(Object instance, PropertyInfo propertyInfo, object[] index)
            : base(propertyInfo.Name, null)
        {
            this.instance = instance;
            this.propertyInfo = propertyInfo;
            this.index = index;
        }

        public override bool CanResetValue(object component)
        {
            return false;
        }

        public override string Category
        {
            get
            {
                return "Properties";
            }
        }

        public override string Description
        {
            get
            {
                return "Shows all properties of the object";
            }
        }

        public override Type ComponentType
        {
            get { return propertyInfo.DeclaringType; }
        }

        public override object GetValue(object component)
        {
            try
            {
                return propertyInfo.GetValue(instance, index);
            }
            catch
            {
            }
            return "Not Available";
        }

        public override bool IsReadOnly
        {
            get { return true; }
        }

        public override Type PropertyType
        {
            get { return propertyInfo.PropertyType; }
        }

        public override void ResetValue(object component)
        {
        }

        public override void SetValue(object component, object value)
        {
        }

        public override bool ShouldSerializeValue(object component)
        {
            return false;
        }

        public override PropertyDescriptorCollection GetChildProperties(object instance, Attribute[] filter)
        {
            if (!propertyInfo.PropertyType.IsValueType)
            {
                return FromObject(propertyInfo.GetValue(instance, index), filter);
            }
            return null;
        }

        public override TypeConverter Converter
        {
            get
            {
                return new ExpandableObjectConverter();
            }
        }

        public override bool IsBrowsable
        {
            get
            {
                return propertyInfo.PropertyType.IsValueType;
            }
        }


        public static PropertyDescriptorCollection FromObject(object instance, Attribute[] attributes)
        {
            Type instanceType = instance.GetType();
            PropertyInfo[] properties = instanceType.GetProperties();
            var propertyDescriptors = properties.Select(propertyInfo => new ReflectionPropertyDescriptor(instance, propertyInfo, null));
            return new PropertyDescriptorCollection(propertyDescriptors.ToArray(), true);
        }
    }
}
