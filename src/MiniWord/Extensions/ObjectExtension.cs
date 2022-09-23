namespace MiniSoftware.Extensions
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Dynamic;

    internal static class ObjectExtension
    {
        internal static Dictionary<string, object> ToDictionary(this object value)
        {
            if (value == null)
                return new Dictionary<string, object>();
            else if (value is IDictionary<string, object> dicStr)
                return (Dictionary<string, object>)dicStr;
            else if (value is ExpandoObject)
                return new Dictionary<string, object>(value as ExpandoObject);

            if (IsArray(value))
                throw new Exception("参数不能为集合类型(The parameter cannot be a collection type)");
            else
            {
                Dictionary<string, object> reuslt = new Dictionary<string, object>();

                PropertyDescriptorCollection props = TypeDescriptor.GetProperties(value);
                foreach (PropertyDescriptor prop in props)
                {

                    object val1 = prop.GetValue(value);

                    //支持第二层，且只支持2层
                    if (IsArray(val1))
                    {
                        List<Dictionary<string, object>> sx = new List<Dictionary<string, object>>();
                        foreach (object val1item in (IEnumerable)val1)
                        {
                            PropertyDescriptorCollection props2 = TypeDescriptor.GetProperties(val1item);
                            Dictionary<string, object> reuslt2 = new Dictionary<string, object>();
                            foreach (PropertyDescriptor prop2 in props2)
                            {
                                object val2 = prop2.GetValue(val1item);
                                if (IsArray(val2) && val2 != null)
                                    throw new Exception("集合类型最多只支持2层(A collection type supports a maximum of two layers)");

                                reuslt2.Add(prop2.Name, val2);
                            }
                            sx.Add(reuslt2);
                        }
                        reuslt.Add(prop.Name, sx);
                    }
                    else
                    {
                        reuslt.Add(prop.Name, val1);
                    }
                }

                return reuslt;
            }
        }

        internal static bool IsArray(object obj)
        {
            return obj is IEnumerable && !(obj is string);
        }

    }
}