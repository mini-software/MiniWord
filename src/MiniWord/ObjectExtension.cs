namespace MiniSoftware
{
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Dynamic;

    internal static class ObjectExtension
	{
        internal static Dictionary<string, object> ToDictionary(this object value)
        {
            if (value is ExpandoObject)
            {
                return new Dictionary<string, object>(value as ExpandoObject);
            }

            Dictionary<string, object> reuslt = new Dictionary<string, object>();

            if (value != null)
            {
                PropertyDescriptorCollection props = TypeDescriptor.GetProperties(value);
                foreach (PropertyDescriptor prop in props)
                {
                    object val = prop.GetValue(value);
                    reuslt.Add(prop.Name, val);
                }
            }

            return reuslt;
        }
    }
}