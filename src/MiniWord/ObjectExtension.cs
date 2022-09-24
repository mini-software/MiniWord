namespace MiniSoftware
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
			else if (value is Dictionary<string, object> dicStr)
				return dicStr;
			else if (value is ExpandoObject)
				return new Dictionary<string, object>(value as ExpandoObject);

			if (IsStrongTypeEnumerable(value))
				throw new Exception("The parameter cannot be a collection type");
			else
			{
				Dictionary<string, object> reuslt = new Dictionary<string, object>();
				PropertyDescriptorCollection props = TypeDescriptor.GetProperties(value);
				foreach (PropertyDescriptor prop in props)
				{
					object val1 = prop.GetValue(value);

					if (IsStrongTypeEnumerable(val1))
					{
						var isValueOrStringType = false; ;
						List<Dictionary<string, object>> sx = new List<Dictionary<string, object>>();
						foreach (object val1item in (IEnumerable)val1)
						{
							if (val1item == null)
							{
								sx.Add(new Dictionary<string, object>());
								continue;
							}
							// not custom type
							if (val1item is string || val1item is DateTime || value.GetType().IsValueType)
							{
								isValueOrStringType = true;
								reuslt.Add(prop.Name, val1);
								break;
							}
							if (val1item is Dictionary<string, object> dicStr)
							{
								sx.Add(dicStr);
								continue;
							}
							else if (val1item is ExpandoObject)
							{
								sx.Add(new Dictionary<string, object>(value as ExpandoObject));
								continue;
							}

							PropertyDescriptorCollection props2 = TypeDescriptor.GetProperties(val1item);
							Dictionary<string, object> reuslt2 = new Dictionary<string, object>();
							foreach (PropertyDescriptor prop2 in props2)
							{
								object val2 = prop2.GetValue(val1item);
								reuslt2.Add(prop2.Name, val2);
							}
							sx.Add(reuslt2);
						}
						if (!isValueOrStringType)
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
		internal static bool IsStrongTypeEnumerable(object obj)
		{
			return obj is IEnumerable && !(obj is string) && !(obj is char[]) && !(obj is string[]);
		}
	}
}