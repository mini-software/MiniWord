namespace MiniSoftware
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Extensions;
    using Utility;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
    using System.Threading.Tasks;
    using System.Threading;

    public static partial class MiniWord
    {
        private static void SaveAsByTemplateImpl(Stream stream, byte[] template, Dictionary<string, object> data)
        {
            var value = data; //TODO: support dynamic and poco value
            byte[] bytes = null;
            using (var ms = new MemoryStream())
            {
                ms.Write(template, 0, template.Length);
                bytes = WriteToByte(data, ms);
            }
            stream.Write(bytes, 0, bytes.Length);
        }

        private static async Task SaveAsByTemplateImplAsync(Stream stream, byte[] template, Dictionary<string, object> data,CancellationToken token)
        {
            byte[] bytes = null;
            using (var ms = new MemoryStream())
            {
                await ms.WriteAsync(template, 0, template.Length, token);
                bytes = WriteToByte(data, ms);
            }
            await stream.WriteAsync(bytes, 0, bytes.Length,token);
        }

        private static byte[] WriteToByte(Dictionary<string, object> value, MemoryStream ms)
        {
            ms.Position = 0;
            using (var docx = WordprocessingDocument.Open(ms, true))
            {
                var hc = docx.MainDocumentPart.HeaderParts.Count();
                var fc = docx.MainDocumentPart.FooterParts.Count();
                for (int i = 0; i < hc; i++)
                    docx.MainDocumentPart.HeaderParts.ElementAt(i).Header.Generate(docx, value);
                for (int i = 0; i < fc; i++)
                    docx.MainDocumentPart.FooterParts.ElementAt(i).Footer.Generate(docx, value);
                docx.MainDocumentPart.Document.Body.Generate(docx, value);
                docx.Save();
            }
            return ms.ToArray();
        }

        private static void Generate(this OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
        {
            // avoid {{tag}} like <t>{</t><t>{</t> 
            //AvoidSplitTagText(xmlElement);
            // avoid {{tag}} like <t>aa{</t><t>{</t>  test in...
            AvoidSplitTagText(xmlElement);

            //Tables
            var tables = xmlElement.Descendants<Table>().ToArray();
            {
                foreach (var table in tables)
                {
                    var trs = table.Descendants<TableRow>().ToArray(); // remember toarray otherwise system will loop OOM;

                    foreach (var tr in trs)
                    {
                        var innerText = tr.InnerText.Replace("{{foreach", "").Replace("endforeach}}", "")
                            .Replace("{{if(", "").Replace(")if", "").Replace("endif}}", "");
                        var matchs = Regex.Matches(innerText, "(?<={{).*?\\..*?(?=}})")
                            .Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value).ToArray();
                        if (matchs.Length > 0)
                        {
                            var listKeys = matchs.Select(s => s.Split('.')[0]).Distinct().ToArray();
                            // TODO:
                            // not support > 2 list in same tr
                            if (listKeys.Length > 2)
                                throw new NotSupportedException("MiniWord doesn't support more than 2 list in same row");
                            var listKey = listKeys[0];
                            if (tags.ContainsKey(listKey) && tags[listKey] is IEnumerable)
                            {
                                var attributeKey = matchs[0].Split('.')[0];
                                var list = tags[listKey] as IEnumerable;

                                foreach (Dictionary<string, object> es in list)
                                {
                                    var dic = es.ToDictionary(e => $"{listKey}.{e.Key}", e => e.Value);
                                    var newTr = tr.CloneNode(true);
                                    ReplaceStatements(newTr, tags: dic);
                                    
                                    ReplaceText(newTr, docx, tags: dic);
                                    //Fix #47 The table should be inserted at the template tag position instead of the last row
                                    if (table.Contains(tr))
                                    {
                                        table.InsertBefore(newTr, tr);
                                    }
                                    else
                                    {
                                        // If it is a nested table, temporarily append it to the end according to the original plan.
                                        table.Append(newTr);
                                    }
                                }
                                tr.Remove();
                            }
                        }
                    }
                }
            }
            
            ReplaceStatements(xmlElement, tags);
            
            ReplaceText(xmlElement, docx, tags);
        }

        private static void AvoidSplitTagText(OpenXmlElement xmlElement)
        {
            var texts = xmlElement.Descendants<Text>().ToList();
            var pool = new List<Text>();
            var sb = new StringBuilder();
            var needAppend = false;
            foreach (var text in texts)
            {
                var clear = false;
                if (text.InnerText.StartsWith("{"))
                {
                    needAppend = true;
                }
                if (needAppend)
                {
                    sb.Append(text.InnerText);
                    pool.Add(text);

                    var s = sb.ToString(); //TODO:
                                           // TODO: check tag exist
                                           // TODO: record tag text if without tag then system need to clear them
                                           // TODO: every {{tag}} one <t>for them</t> and add text before first text and copy first one and remove {{, tagname, }}
                    
                    const string foreachTag = "{{foreach";
                    const string endForeachTag = "endforeach}}";
                    const string ifTag = "{{if";
                    const string endifTag = "endif}}";
                    const string tagStart = "{{";
                    const string tagEnd = "}}";
                    
                    var foreachTagContains = s.Split(new []{foreachTag}, StringSplitOptions.None).Length - 1 ==
                                             s.Split(new []{endForeachTag}, StringSplitOptions.None).Length - 1;
                    var ifTagContains = s.Split(new []{ifTag}, StringSplitOptions.None).Length - 1 ==
                                        s.Split(new []{endifTag}, StringSplitOptions.None).Length - 1;
                    var tagContains = s.StartsWith(tagStart) && s.Contains(tagEnd);
                    
                    if (foreachTagContains && ifTagContains && tagContains)
                    {
                        if (sb.Length <= 1000) // avoid too big tag
                        {
                            var first = pool.First();
                            var newText = first.Clone() as Text;
                            newText.Text = s;
                            first.Parent.InsertBefore(newText, first);
                            foreach (var t in pool)
                            {
                                t.Text = "";
                            }
                        }
                        clear = true;
                    }
                }

                if (clear)
                {
                    sb.Clear();
                    pool.Clear();
                    needAppend = false;
                }
            }
        }

        private static void AvoidSplitTagText(OpenXmlElement xmlElement, IEnumerable<string> txt)
        {
            foreach (var paragraph in xmlElement.Descendants<Paragraph>())
            {
                foreach (var continuousString in paragraph.GetContinuousString())
                {
                    foreach (var text in txt.Where(o => continuousString.Item1.Contains(o)))
                    {
                        continuousString.Item3.TrimStringToInContinuousString(text);
                    }
                }
            }
        }

        private static List<string> GetReplaceKeys(Dictionary<string, object> tags)
        {
            var keys = new List<string>();
            foreach (var item in tags)
            {
                if (item.Value.IsStrongTypeEnumerable())
                {
                    foreach (var item2 in (IEnumerable)item.Value)
                    {
                        if (item2 is Dictionary<string, object> dic)
                        {
                            foreach (var item3 in dic.Keys)
                            {
                                keys.Add("{{" + item.Key + "." + item3 + "}}");
                            }
                        }
                        break;
                    }
                }
                else
                {
                    keys.Add("{{" + item.Key + "}}");
                }
            }
            return keys;
        }

        private static bool EvaluateStatement(string tagValue, string comparisonOperator, string value)
        {
            var checkStatement = false;

            var tagValueEvaluation = EvaluateValue(tagValue);
            
            switch (tagValueEvaluation)
            {
                case double dtg:
                    if(double.TryParse(value, out var doubleNumber)) {
                        switch (comparisonOperator)
                        {
                            case "==":
                                checkStatement = dtg.Equals(doubleNumber);
                                break;
                            case "!=":
                                checkStatement = !dtg.Equals(doubleNumber);
                                break;
                            case ">":
                                checkStatement = dtg > doubleNumber;
                                break;
                            case "<":
                                checkStatement = dtg < doubleNumber;
                                break;
                            case ">=":
                                checkStatement = dtg >= doubleNumber;
                                break;
                            case "<=":
                                checkStatement = dtg <= doubleNumber;
                                break;
                        }
                    } else {
                        checkStatement = comparisonOperator == "!=";                            
                    }

                    break;
                case int itg:
                    if(int.TryParse(value, out var intNumber))
                    {
                        switch (comparisonOperator)
                        {
                            case "==":
                                checkStatement = itg.Equals(intNumber);
                                break;
                            case "!=":
                                checkStatement = !itg.Equals(intNumber);
                                break;
                            case ">":
                                checkStatement = itg > intNumber;
                                break;
                            case "<":
                                checkStatement = itg < intNumber;
                                break;
                            case ">=":
                                checkStatement = itg >= intNumber;
                                break;
                            case "<=":
                                checkStatement = itg <= intNumber;
                                break;
                        }
                    } else {
                        checkStatement = comparisonOperator == "!=";                            
                    }

                    break;
                case DateTime dttg:
                    if(DateTime.TryParse(value, out var date))
                    {
                        switch (comparisonOperator)
                        {
                            case "==":
                                checkStatement = dttg.Equals(date);
                                break;
                            case "!=":
                                checkStatement = !dttg.Equals(date);
                                break;
                            case ">":
                                checkStatement = dttg > date;
                                break;
                            case "<":
                                checkStatement = dttg < date;
                                break;
                            case ">=":
                                checkStatement = dttg >= date;
                                break;
                            case "<=":
                                checkStatement = dttg <= date;
                                break;
                        }
                    } else {
                        checkStatement = comparisonOperator == "!=";
                    }

                    break;
                case string stg:
                    switch (comparisonOperator)
                    {
                        case "==":
                            checkStatement = stg == value;
                            break;
                        case "!=":
                            checkStatement = stg != value;
                            break;
                    }

                    break;
                case bool btg when bool.TryParse(value, out var boolean):
                    switch (comparisonOperator)
                    {
                        case "==":
                            checkStatement = btg != boolean;
                            break;
                        case "!=":
                            checkStatement = btg == boolean;
                            break;
                    }
                    
                    break;
            }

            return checkStatement;
        }

        private static object EvaluateValue(string value)
        {
            if (double.TryParse(value, out var doubleNumber))
                return doubleNumber;
            else if (int.TryParse(value, out var intNumber))
                return intNumber;
            else if (DateTime.TryParse(value, out var date))
                return date;
            
            return value;
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

                            if (!isMatch && tag.Value is List<MiniWordForeach> forTags)
                            {
                                if (forTags.Any(forTag => forTag.Value.Keys.Any(dictKey =>
                                    {
                                        return t.Text.Contains($@"{{{{{tag.Key}.{dictKey}}}}}");
                                    })))
                                {
                                    isMatch = true;
                                }
                            }

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
                                        newT.Text = EvaluateIfStatement(newT.Text);
                                        run.Append(newT);
                                        currentT = newT;
                                    }
                                    t.Remove();
                                }
                                else if (tag.Value is List<MiniWordForeach> vs)
                                {
                                    var currentT = t;
                                    var generatedText = new Text();
                                    currentT.Text = currentT.Text.Replace(@"{{foreach", "").Replace(@"endforeach}}", "");
                                    
                                    var newTexts = new Dictionary<int, string>();
                                    for (var i = 0; i < vs.Count; i++)
                                    {
                                        var newT = t.CloneNode(true) as Text;

                                        foreach (var vv in vs[i].Value)
                                        {
                                            newT.Text = newT.Text.Replace("{{" + tag.Key + "." + vv.Key + "}}", vv.Value.ToString());
                                        }
                                        
                                        newT.Text = EvaluateIfStatement(newT.Text);
                                        
                                        if(!string.IsNullOrEmpty(newT.Text))
                                            newTexts.Add(i, newT.Text);
                                    }

                                    for (var i = 0; i < newTexts.Count; i++)
                                    {
                                        var dict = newTexts.ElementAt(i);
                                        generatedText.Text += dict.Value;
                                        
                                        if (i != newTexts.Count - 1)
                                        {
                                            generatedText.Text += vs[dict.Key].Separator;
                                        }
                                    }

                                    run.Append(generatedText);
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
                                else if (IsHyperLink(tag.Value))
                                {
                                    AddHyperLink(docx, run, tag.Value);
                                    t.Remove();
                                }
                                else if (tag.Value is MiniWordColorText || tag.Value is MiniWordColorText[])
                                {
                                    var colorText = tag.Value is MiniWordColorText
                                        ? AddColorText(new[] { (MiniWordColorText)tag.Value })
                                        : AddColorText((MiniWordColorText[])tag.Value);
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

                        t.Text = EvaluateIfStatement(t.Text);

                        // add breakline
                        {
                            var newText = t.Text;
                            var splits = Regex.Split(newText, "(<[a-zA-Z/].*?>|\n|\r\n)").Where(o => o != "\n" && o != "\r\n");
                            var currentT = t;
                            var isFirst = true;
                            if (splits.Count() > 1)
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
        
        private static void ReplaceStatements(OpenXmlElement xmlElement, Dictionary<string, object> tags)
        {
            var paragraphs = xmlElement.Descendants<Paragraph>().ToList();

            while (paragraphs.Any(s => s.InnerText.Contains("@if")))
            {
                var ifIndex = paragraphs.FindIndex(0, s => s.InnerText.Contains("@if"));
                var endIfFinalIndex = paragraphs.FindIndex(ifIndex, s => s.InnerText.Contains("@endif"));

                var statement = paragraphs[ifIndex].InnerText.Split(' ');

                var tagValue = tags[statement[1]] ?? "NULL";

                var checkStatement = statement.Length == 4 ? EvaluateStatement(tagValue.ToString(), statement[2], statement[3]) : !bool.Parse(tagValue.ToString());

                if (!checkStatement)
                {
                    for (int i = ifIndex + 1; i <= endIfFinalIndex - 1; i++)
                    {
                        paragraphs[i].Remove();
                    }
                }

                paragraphs[ifIndex].Remove();
                paragraphs[endIfFinalIndex].Remove();

                paragraphs = xmlElement.Descendants<Paragraph>().ToList();
            }
        }

        private static string EvaluateIfStatement(string text)
        {
            const string ifStartTag = "{{if(";
            const string ifEndTag = ")if";
            const string endIfTag = "endif}}";
            
            while (text.Contains(ifStartTag)) 
            {
                var ifIndex = text.IndexOf(ifStartTag, StringComparison.Ordinal);
                var ifEndIndex = text.IndexOf(")if", ifIndex, StringComparison.Ordinal);
                            
                var statement = text.Substring(ifIndex + ifStartTag.Length, ifEndIndex - (ifIndex + ifStartTag.Length)).Split(',');
                            
                var checkStatement = EvaluateStatement(statement[0], statement[1], statement[2]);

                if (checkStatement)
                {
                    text = text.Remove(ifIndex, ifEndIndex - ifIndex + ifEndTag.Length);
                    var endIfFinalIndex = text.IndexOf(endIfTag, StringComparison.Ordinal);
                    text = text.Remove(endIfFinalIndex, endIfTag.Length);
                }
                else
                {
                    var endIfFinalIndex = text.IndexOf(endIfTag, StringComparison.Ordinal);
                    text = text.Remove(ifIndex, endIfFinalIndex - ifIndex + endIfTag.Length);
                }
            }

            return text;
        }

        private static bool IsHyperLink(object value)
        {
            return value is MiniWordHyperLink ||
                    value is IEnumerable<MiniWordHyperLink>;
        }

        private static void AddHyperLink(WordprocessingDocument docx, Run run, object value)
        {
            List<MiniWordHyperLink> links = new List<MiniWordHyperLink>();

            if (value is MiniWordHyperLink)
            {
                links.Add((MiniWordHyperLink)value);
            }
            else
            {
                links.AddRange((IEnumerable<MiniWordHyperLink>)value);
            }

            foreach (var linkInfo in links)
            {
                var mainPart = docx.MainDocumentPart;
                var hyperlink = GetHyperLink(mainPart, linkInfo);
                run.Append(hyperlink);
                run.Append(new Break());
            }
        }

        private static Hyperlink GetHyperLink(MainDocumentPart mainPart, MiniWordHyperLink linkInfo)
        {
            var hr = mainPart.AddHyperlinkRelationship(new Uri(linkInfo.Url), true);
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
        private static RunProperties AddColorText(MiniWordColorText[] miniWordColorTextArray)
        {
            RunProperties runPro = new RunProperties();
            foreach (var miniWordColorText in miniWordColorTextArray)
            {
                Text text = new Text(miniWordColorText.Text);
                Color color = new Color() { Val = miniWordColorText.FontColor?.Replace("#", "") };
                Shading shading = new Shading() { Fill = miniWordColorText.HighlightColor?.Replace("#", "") };
                runPro.Append(shading);
                runPro.Append(color);
                runPro.Append(text);
            }

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
                                                        $"{{{Guid.NewGuid().ToString("n")}}}"
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

        private static async Task<byte[]> GetByteAsync(string path)
        {
            using (var st = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, true))
            using (var ms = new MemoryStream())
            {
                await st.CopyToAsync(ms);
                return ms.ToArray();
            }
        }
    }
}