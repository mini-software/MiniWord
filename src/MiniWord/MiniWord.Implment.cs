namespace MiniSoftware
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using MiniSoftware.Utility;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

    public static partial class MiniWord
    {
        private static void SaveAsByTemplateImpl(Stream stream, byte[] template, Dictionary<string, object> data)
        {
            var value = data; //TODO: support dynamic and poco value
            byte[] bytes = null;
            using (var ms = new MemoryStream())
            {
                ms.Write(template, 0, template.Length);
                ms.Position = 0;
                using (var docx = WordprocessingDocument.Open(ms, true))
                {
                    foreach (var hdr in docx.MainDocumentPart.HeaderParts)
                        hdr.Header.Generate(docx, value);
                    foreach (var ftr in docx.MainDocumentPart.FooterParts)
                        ftr.Footer.Generate(docx, value);
                    docx.MainDocumentPart.Document.Body.Generate(docx, value);
                    docx.Save();
                }
                bytes = ms.ToArray();
            }
            stream.Write(bytes, 0, bytes.Length);
        }
        private static void Generate(this OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
        {

            // Avoid not standard string format e.g. {{<t>tag</t>}}
            //foreach (var tag in tags)
            //{
            //    var regexStr = string.Concat(@"((\{\{(?:(?!\{\{|}}).)*>)|\{\{)", tag.Key, @"(}}|<.*?}})");

            //    xmlElement.InnerXml = Regex.Replace(xmlElement.InnerXml, regexStr, $"{{{{{tag.Key?.ToString()}}}}}", RegexOptions.Singleline | RegexOptions.IgnorePatternWhitespace | RegexOptions.CultureInvariant);
            //}
            // Avoid not standard string format e.g. {{<t>tag</t>}}
            {
                var paragraphs = xmlElement.Descendants<Paragraph>().ToArray();
                foreach (var p in paragraphs)
                {
                    var runs = p.Descendants<Run>().ToArray();
                    var isMatch = tags.Any(tag =>
                    {
                        var b = p.InnerText.Contains($"{{{{{tag.Key}}}}}");
                        if (!b && tag.Value is IEnumerable)
                        {
                            foreach (var item in tag.Value as IEnumerable)
                            {
                                if (item is Dictionary<string, object>)
                                {
                                    foreach (var dic in item as Dictionary<string, object>)
                                    {
                                        b = p.InnerText.Contains($"{{{{{tag.Key}.{dic.Key}}}}}");
                                        if (b)
                                            break;
                                    }
                                }
                                break;
                            }
                        }
                        return b;
                    });
                    if (isMatch)
                    {
                        var newText = p.InnerText?.ToString();
                        foreach (var run in runs.Skip(1))
                            run.RemoveAllChildren<Text>();
                        if (runs.Length > 0)
                        {
                            var texts = runs[0].Descendants<Text>().ToArray();
                            if (texts.Length > 0)
                            {
                                foreach (var text in texts.Skip(1))
                                    text.RemoveAllChildren();
                                texts[0].Text = newText;
                            }
                        }
                    }
                }
            }
            //Tables
            var tables = xmlElement.Descendants<Table>().ToArray();
            {
                foreach (var table in tables)
                {
                    var trs = table.Descendants<TableRow>().ToArray(); // remember toarray or system will loop OOM;

                    foreach (var tr in trs)
                    {
                        
                        var matchs = (Regex.Matches(tr.InnerText, "(?<={{).*?\\..*?(?=}})")
                            .Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value)).ToArray();
                        if (matchs.Length > 0)
                        {
                            var listKeys = matchs.Select(s => s.Split('.')[0]).Distinct().ToArray();
                            // TODO:
                            // not support > 1 list in same tr
                            if (listKeys.Length > 1)
                                throw new NotSupportedException("MiniWord doesn't support more than 1 list in same row");
                            var listKey = listKeys[0];
                            if (tags.ContainsKey(listKey) && tags[listKey] is IEnumerable)
                            {
                                var attributeKey = matchs[0].Split('.')[0];
                                var list = tags[listKey] as IEnumerable;

                                foreach (Dictionary<string, object> es in list)
                                {
                                    var dic = new Dictionary<string, object>(); //TODO: optimize

                                    var newTr = tr.CloneNode(true);
                                    foreach (var e in es)
                                    {
                                        var dicKey = $"{listKey}.{e.Key}";
                                        dic.Add(dicKey, e.Value);
                                    }

                                    ReplaceText(newTr, docx, tags : dic);
                                    table.Append(newTr);
                                }
                                tr.Remove();
                            }
                        }
                    }
                }
            }

            ReplaceText(xmlElement, docx, tags);
        }

        private static void ReplaceText(OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
        {
            var paragraphs = xmlElement.Descendants<Paragraph>().ToArray();
            foreach (var p in paragraphs)
            {
                var runs = p.Descendants<Run>().ToArray();

                foreach (var run in runs)
                {
                    var texts = run.Descendants<Text>().ToArray();
                    if (texts.Length == 0)
                        continue;
                    foreach (Text t in texts)
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
                                        if (isFirst)
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
                                else if(tag.Value is MiniWorHyperLink){
                                    var mainPart = docx.MainDocumentPart;
                                    var linkInfo = (MiniWorHyperLink)tag.Value;
                                    var hyperlink = GetHyperLink(mainPart,linkInfo);
                                    run.Append(hyperlink);
                                    t.Remove();
                                }
                                else if (tag.Value is MiniWordColorText)
                                {
                                    var miniWordColorText = (MiniWordColorText)tag.Value;
                                    var colorText = AddColorText(miniWordColorText);
                                    run.Append(colorText);
                                    t.Remove();
                                }
                                else
                                {
                                    var newText = string.Empty;
                                    if (tag.Value is DateTime)
                                    {
                                        newText = ((DateTime)tag.Value).ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    else
                                    {
                                        newText = tag.Value?.ToString();
                                    }
                                    t.Text = t.Text.Replace($"{{{{{tag.Key}}}}}", newText);
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

        private static Hyperlink GetHyperLink(MainDocumentPart mainPart,MiniWorHyperLink linkInfo)
        {
            var hr = mainPart.AddHyperlinkRelationship(new Uri(linkInfo.Url),true);
            Hyperlink xmlHyperLink = new Hyperlink(new RunProperties(
                    new RunStyle { Val = "Hyperlink", },
                    new Underline { Val = linkInfo.UnderLineValue },
                    new Color { ThemeColor = ThemeColorValues.Hyperlink }),
                new Text(linkInfo.Text)
                )
            {
                DocLocation = linkInfo.Url,
                Id = hr.Id,
                TargetFrame = linkInfo.GetTargetFrame()
            };
            return xmlHyperLink; 
        }
        private static RunProperties AddColorText(MiniWordColorText miniWordColorText)
        {

            RunProperties runPro = new RunProperties();
            Text text = new Text(miniWordColorText.Text);
            Color color = new Color() { Val = miniWordColorText.ForeColor?.Replace("#", "") };
            Shading shading = new Shading() { Fill = miniWordColorText.BackColor?.Replace("#", "") };
            runPro.Append(shading);
            runPro.Append(color);
            runPro.Append(text);

            return runPro;
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
        private static byte[] GetBytes(string path)
        {
            using (var st = Helpers.OpenSharedRead(path))
            using (var ms = new MemoryStream())
            {
                st.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}