namespace MiniSoftware
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

    public static class MiniWord
	{
		static void ReplaceTag(this OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
		{
			
			// Avoid not standard string format e.g. {{<t>tag</t>}}
			foreach (var tag in tags)
            {
				var regexStr = string.Concat(@"((\{\{(?:(?!\{\{|}}).)*>)|\{\{)", tag.Key, @"(}}|<.*?}})");

				xmlElement.InnerXml = Regex.Replace(xmlElement.InnerXml, regexStr, $"{{{{{tag.Key?.ToString()}}}}}", RegexOptions.Singleline | RegexOptions.IgnorePatternWhitespace | RegexOptions.CultureInvariant);
			}
			//return;
			var paragraphs = xmlElement.Descendants<Paragraph>();
			foreach (var p in paragraphs)
			{
				var runs = p.Descendants<Run>();
				foreach (var run in runs)
                {
                    foreach (Text t in run.Descendants<Text>())
                    {
						foreach (var tag in tags)
                        {
							var isMatch = t.Text.Contains($"{{{{{tag.Key}}}}}");
							if (isMatch)
                            {
								if (tag.Value is string[] || tag.Value is IList<string> || tag.Value is List<string>)
                                {
									var vs = tag.Value as IEnumerable;
									var currentT = t;
									var isFirst = true;
                                    foreach (var v in vs)
                                    {	
										var newT = t.CloneNode(true) as Text;
										newT.Text = t.Text.Replace($"{{{{{tag.Key}}}}}", v?.ToString());
										if(isFirst)
											isFirst = false;
										else
											run.Append(new Break());
										run.Append(newT);
										currentT = newT;
									}
                                    t.Remove();
                                }
                                else if (tag.Value is MiniWordPicture)
                                {
									var pic = (MiniWordPicture)tag.Value;
									byte[] l_Data = null;
									if (pic.Path != null)
                                    {
										l_Data = File.ReadAllBytes(pic.Path);
									}
									if (pic.Bytes != null)
									{
										l_Data = pic.Bytes;
									}

									var mainPart = docx.MainDocumentPart;

									var imagePart = mainPart.AddImagePart(pic.GetImagePartType);
                                    using (var stream = new MemoryStream(l_Data))
                                    {
                                        imagePart.FeedData(stream);
                                        AddPicture(run, mainPart.GetIdOfPart(imagePart), pic);

									}
									t.Remove();
                                }
                                else
                                {
									t.Text = t.Text.Replace($"{{{{{tag.Key}}}}}", tag.Value?.ToString());
								}
							}
						}

						// add breakline
                        {
							var newText = t.Text;
							var splits = Regex.Split(newText, "(<[a-zA-Z/].*?>|\n)");
							var currentT = t;
							var isFirst = true;
                            if (splits.Length > 1)
                            {
								foreach (var v in splits)
								{
									var newT = t.CloneNode(true) as Text;
									newT.Text = v?.ToString();
									if (isFirst)
										isFirst = false;
									else
										run.Append(new Break());
									run.Append(newT);
									currentT = newT;
								}
								t.Remove();
							}
						}
					}
                }
			}
		}

		private static void AddPicture(OpenXmlElement appendElement, string relationshipId, MiniWordPicture pic)
		{
			// Define the reference of the image.
			var element =
				 new Drawing(
					 new DW.Inline(
						 new DW.Extent() { Cx = pic.Cx, Cy = pic.Cy },
						 new DW.EffectExtent()
						 {
							 LeftEdge = 0L,
							 TopEdge = 0L,
							 RightEdge = 0L,
							 BottomEdge = 0L
						 },
						 new DW.DocProperties()
						 {
							 Id = (UInt32Value)1U,
							 Name = $"Picture {Guid.NewGuid().ToString()}"
						 },
						 new DW.NonVisualGraphicFrameDrawingProperties(
							 new A.GraphicFrameLocks() { NoChangeAspect = true }),
						 new A.Graphic(
							 new A.GraphicData(
								 new PIC.Picture(
									 new PIC.NonVisualPictureProperties(
										 new PIC.NonVisualDrawingProperties()
										 {
											 Id = (UInt32Value)0U,
											 Name = $"Image {Guid.NewGuid().ToString()}.{pic.Extension}"
										 },
										 new PIC.NonVisualPictureDrawingProperties()),
									 new PIC.BlipFill(
										 new A.Blip(
											 new A.BlipExtensionList(
												 new A.BlipExtension()
												 {
													 Uri =
														$"{{{ Guid.NewGuid().ToString("n")}}}"
												 })
										 )
										 {
											 Embed = relationshipId,
											 CompressionState =
											 A.BlipCompressionValues.Print
										 },
										 new A.Stretch(
											 new A.FillRectangle())),
									 new PIC.ShapeProperties(
										 new A.Transform2D(
											 new A.Offset() { X = 0L, Y = 0L },
											 new A.Extents() { Cx = pic.Cx, Cy = pic.Cy }),
										 new A.PresetGeometry(
											 new A.AdjustValueList()
										 )
										 { Preset = A.ShapeTypeValues.Rectangle }))
							 )
							 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
					 )
					 {
						 DistanceFromTop = (UInt32Value)0U,
						 DistanceFromBottom = (UInt32Value)0U,
						 DistanceFromLeft = (UInt32Value)0U,
						 DistanceFromRight = (UInt32Value)0U,
						 EditId = "50D07946"
				 });
			appendElement.Append((element));
		}
		public static void SaveAsByTemplate(string path, string templatePath, Dictionary<string,object> value)
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
		private static byte[] GetBytes(string path)
		{
			using (var st = Helpers.OpenSharedRead(path))
			using (var ms = new MemoryStream())
			{
				st.CopyTo(ms);
				return ms.ToArray();
			}
		}
		private static void SaveAsByTemplateImpl(Stream stream, byte[] template, Dictionary<string, object> data)
		{
			var value = data ; //TODO: support dynamic and poco value
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
}