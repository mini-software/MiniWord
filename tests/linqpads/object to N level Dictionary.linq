<Query Kind="Program">
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>DocumentFormat.OpenXml</Namespace>
  <Namespace>DocumentFormat.OpenXml.Packaging</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.ComponentModel</Namespace>
  <Namespace>System.Dynamic</Namespace>
</Query>

void Main()
{
{
		var value = new
		{
			TripHs = new List<Dictionary<string, object>>
				{
					new Dictionary<string, object>
					{
						{ "sDate",DateTime.Parse("2022-09-08 08:30:00")},
						{ "eDate",DateTime.Parse("2022-09-08 15:00:00")},
						{ "How","Discussion requirement part1"},
						{ "Photo",new MiniWordPicture() { Path = "DemoExpenseMeeting02.png", Width = 160, Height = 90 }},
					},
					new Dictionary<string, object>
					{
						{ "sDate",DateTime.Parse("2022-09-09 08:30:00")},
						{ "eDate",DateTime.Parse("2022-09-09 17:00:00")},
						{ "How","Discussion requirement part2 and development"},
						{ "Photo",new MiniWordPicture() { Path = "DemoExpenseMeeting02.png", Width = 160, Height = 90 }},
					},
				}
		};
		var data = value.ToDictionary();
		Console.WriteLine(data);
	}
	{
		var value = new
		{
			managers = new List<string> { "Jack", "Alan" },
			employees = new List<DateTime?> { null, new DateTime(2011,2,24) },
		};
		var data = value.ToDictionary();
		Console.WriteLine(data);
	}

	// How 2 level object convert to 2 level Dictionary<string,object> this object can be Enumerable<Dictionary<string,object>> or object
	var department = new
	{
		ID = "S001",
		Name = "HR",
		Users = new[]{
			new {ID="E001",Name="Jack"},
			new {ID="E002",Name="Terry"},
			new {ID="E003",Name="Jassie"},
		},
		ChildDepartments = new[] {"D004","D005","D006"},
		Times = new[] {new DateTime(2022,10,15),new DateTime(2022,11,08),new DateTime(2022,12,25)},
	};
	//var json = JsonConvert.SerializeObject(department);
	//Console.WriteLine(json); //{"ID":"S001","Name":"HR","Users":[{"ID":"E001","Name":"Jack"},{"ID":"E002","Name":"Terry"},{"ID":"E003","Name":"Jassie"}]}


	var result = ObjectExtension.ToDictionary(department);
	//Console.WriteLine(result);

	var expectedResult = new Dictionary<string, object>
	{
		["ID"] = "S001",
		["Name"] = "HR",
		["Users"] = new List<Dictionary<string, object>>{
			new Dictionary<string,object>{["ID"]="E001",["Name"]="Jack"},
			new Dictionary<string,object>{["ID"]="E002",["Name"]="Terry"},
			new Dictionary<string,object>{["ID"]="E003",["Name"]="Jassie"},
		},
		["ChildDepartments"] = new[] {"D004","D005","D006"},
		["Times"] = new[] {new DateTime(2022,10,15),new DateTime(2022,11,08),new DateTime(2022,12,25)},
	};
	//Console.WriteLine(expectedResult);


	//var data = JsonConvert.DeserializeObject<Dictionary<string,Dictionary<string,object>>>(json);
	//Console.WriteLine(data);

}

// You can define other methods, fields, classes and namespaces here
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
					var isValueOrStringType = false;;
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
							isValueOrStringType=true;
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
					if(!isValueOrStringType)
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



public class MiniWordPicture
{


	public string Path { get; set; }
	private string _extension;
	public string Extension
	{
		get
		{
			if (Path != null)
				return System.IO.Path.GetExtension(Path).ToUpperInvariant().Replace(".", "");
			else
			{
				return _extension.ToUpper();
			}
		}
		set { _extension = value; }
	}
	internal ImagePartType GetImagePartType
	{
		get
		{
			switch (Extension.ToLower())
			{
				case "bmp": return ImagePartType.Bmp;
				case "emf": return ImagePartType.Emf;
				case "ico": return ImagePartType.Icon;
				case "jpg": return ImagePartType.Jpeg;
				case "jpeg": return ImagePartType.Jpeg;
				case "pcx": return ImagePartType.Pcx;
				case "png": return ImagePartType.Png;
				case "svg": return ImagePartType.Svg;
				case "tiff": return ImagePartType.Tiff;
				case "wmf": return ImagePartType.Wmf;
				default:
					throw new NotSupportedException($"{_extension} is not supported");
			}
		}
	}

	public byte[] Bytes { get; set; }
	/// <summary>
	/// Unit is Pixel
	/// </summary>
	public Int64Value Width { get; set; } = 400;
	internal Int64Value Cx { get { return Width * 9525; } }
	/// <summary>
	/// Unit is Pixel
	/// </summary>
	public Int64Value Height { get; set; } = 400;
	//format resource from http://openxmltrix.blogspot.com/2011/04/updating-images-in-image-placeholde-and.html
	internal Int64Value Cy { get { return Height * 9525; } }
}