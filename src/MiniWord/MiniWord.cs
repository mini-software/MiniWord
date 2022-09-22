namespace MiniSoftware
{
	using DocumentFormat.OpenXml.Office2013.Excel;
	using System.Collections.Generic;
	using System.ComponentModel;
	using System.Dynamic;
	using System.IO;
	using System.Linq.Expressions;

	public static partial class MiniWord
	{
		public static void SaveAsByTemplate(string path, string templatePath, Dictionary<string, object> value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templatePath, value);
		}

		public static void SaveAsByTemplate(string path, string templatePath, object value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templatePath, value.ToDictionary());
		}

		public static void SaveAsByTemplate(string path, byte[] templateBytes, Dictionary<string, object> value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templateBytes, value);
		}

		public static void SaveAsByTemplate(string path, byte[] templateBytes, object value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templateBytes, value.ToDictionary());
		}

		public static void SaveAsByTemplate(this Stream stream, string templatePath, Dictionary<string, object> value)
		{
			SaveAsByTemplateImpl(stream, GetBytes(templatePath), value);
		}

        public static void SaveAsByTemplate(this Stream stream, string templatePath, object value)
        {
            SaveAsByTemplateImpl(stream, GetBytes(templatePath), value.ToDictionary());
        }

        public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, Dictionary<string, object> value)
		{
			SaveAsByTemplateImpl(stream, templateBytes, value);
		}

		public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object value)
		{

			SaveAsByTemplateImpl(stream, templateBytes, value.ToDictionary());
		}
    }
}