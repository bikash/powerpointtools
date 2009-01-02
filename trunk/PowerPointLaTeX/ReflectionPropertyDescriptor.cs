using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PowerPointLaTeX
{
    class ReflectionTypeConverter : TypeConverter
    {
        private Type componentType;

        public ReflectionTypeConverter(Type componentType)
        {
            this.componentType = componentType;
        }

        public override bool GetPropertiesSupported(ITypeDescriptorContext context)
        {
            return !componentType.IsValueType;
        }

        public override PropertyDescriptorCollection GetProperties(ITypeDescriptorContext context, object value, Attribute[] attributes)
        {
            PropertyInfo[] properties = componentType.GetProperties();
            PropertyDescriptor[] propertyDescriptors = properties.Select(propertyInfo => new PropertyDescriptor(componentType, propertyInfo)).ToArray();
            return new PropertyDescriptorCollection(propertyDescriptors, true);
        }

        class PropertyDescriptor : SimplePropertyDescriptor
        {
            private PropertyInfo propertyInfo;

            public PropertyDescriptor(Type componentType, PropertyInfo propertyInfo)
                : base(componentType, propertyInfo.Name, propertyInfo.PropertyType)
            {
                this.propertyInfo = propertyInfo;
            }

            public override object GetValue(object component)
            {
                object value = "Not Available";
                try
                {
                    value = propertyInfo.GetValue(component, null);
                }
                catch
                {
                }
                return value;
            }

            public override void SetValue(object component, object value)
            {
            }

            public override string Category
            {
                get
                {
                    return "Properties";
                }
            }

            public override TypeConverter Converter
            {
                get
                {
                    return new ReflectionTypeConverter(PropertyType);
                }
            }

            public override PropertyDescriptorCollection GetChildProperties(object instance, Attribute[] filter)
            {
                return Converter.GetProperties(null, GetValue(instance), filter);
            }
        }
    }

}
