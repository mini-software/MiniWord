namespace MiniSoftware
{
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;

    public static class MiniWord
	{
		static void ReplaceTag(this OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
		{
			var paragraphs = xmlElement.Descendants<Paragraph>();
			foreach (var p in paragraphs)
			{
				var innerXmlSb = p.InnerXml;
				foreach (var tag in tags)
					innerXmlSb = Regex.Replace(innerXmlSb, @"((\{\{(?:(?!\{\{|}}).)*>)|\{\{)" + tag.Key + @"(}}|<.*?}})", tags[tag.Key]?.ToString(), RegexOptions.Singleline | RegexOptions.IgnorePatternWhitespace | RegexOptions.CultureInvariant);
				p.InnerXml = innerXmlSb;
			}
		}
		public static void SaveAsByTemplate(string path, string templatePath, object value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templatePath, value);
		}

		public static void SaveAsByTemplate(string path, byte[] templateBytes, object value)
		{
			using (var stream = File.Create(path))
				SaveAsByTemplate(stream, templateBytes, value);
		}

		public static void SaveAsByTemplate(this Stream stream, string templatePath, object value)
		{
			SaveAsByTemplateImpl(stream, GetBytes(templatePath), value);
		}

		public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object value)
		{
			SaveAsByTemplateImpl(stream, templateBytes, value);
		}
		private static byte[] GetBytes(string path)
		{
			using (var st = FileHelper.OpenSharedRead(path))
			using (var ms = new MemoryStream())
			{
				st.CopyTo(ms);
				return ms.ToArray();
			}
		}
		private static void SaveAsByTemplateImpl(Stream stream, byte[] template, object data)
		{
			var value = data as Dictionary<string, object>; //TODO:
			byte[] bytes = null;
			using (var ms = new MemoryStream())
			{
				ms.Write(template,0, template.Length);
				ms.Position = 0;
				using (var docx = WordprocessingDocument.Open(ms, true))
				{
					foreach (var hdr in docx.MainDocumentPart.HeaderParts)
						hdr.Header.ReplaceTag(docx, value);
					foreach (var ftr in docx.MainDocumentPart.FooterParts)
						ftr.Footer.ReplaceTag(docx, value);
					docx.MainDocumentPart.Document.Body.ReplaceTag(docx, value);
					docx.Save();
				}
				bytes = ms.ToArray();
			}
			stream.Write(bytes,0, bytes.Length);
		}
	}

	internal static partial class FileHelper
	{
		public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
	}
}