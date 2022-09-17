namespace MiniSoftware
{
    using System.Collections.Generic;
    using System.IO;

    public static partial class MiniWord
	{
		public static void SaveAsByTemplate(string path, string templatePath, Dictionary<string, object> value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templatePath, value);
		}

		public static void SaveAsByTemplate(string path, byte[] templateBytes, Dictionary<string, object> value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templateBytes, value);
		}

		public static void SaveAsByTemplate(this Stream stream, string templatePath, Dictionary<string, object> value)
		{
			SaveAsByTemplateImpl(stream, GetBytes(templatePath), value);
		}

		public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, Dictionary<string, object> value)
		{
			SaveAsByTemplateImpl(stream, templateBytes, value);
		}
	}
}