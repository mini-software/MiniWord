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
    using System.Xml;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Drawing.Charts;

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
                    var hc = docx.MainDocumentPart.HeaderParts.Count();
                    var fc = docx.MainDocumentPart.FooterParts.Count();
                    for (int i = 0; i < hc; i++)
                        docx.MainDocumentPart.HeaderParts.ElementAt(i).Header.Generate(docx, value);
                    for (int i = 0; i < fc; i++)
                        docx.MainDocumentPart.FooterParts.ElementAt(i).Footer.Generate(docx, value);
                    docx.MainDocumentPart.Document.Body.Generate(docx, value);
                    docx.Save();
                }
                bytes = ms.ToArray();
            }
            stream.Write(bytes, 0, bytes.Length);
        }
        private static void Generate(this OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
        {
            // avoid {{tag}} like <t>{</t><t>{</t> 
            //AvoidSplitTagText(xmlElement);
            // avoid {{tag}} like <t>aa{</t><t>{</t>  test in...
            AvoidSplitTagText(xmlElement);

            // @foreach循环体
            ReplaceForeachStatements(xmlElement,docx,tags);

            //Tables
            // 忽略table中没有占位符“{{}}”的表格
            var tables = xmlElement.Descendants<Table>().Where(t => t.InnerText.Contains("{{")).ToArray();
            {
                foreach (var table in tables)
                {
                    GenerateTable(table,docx,tags);
                }
            }

            ReplaceIfStatements(xmlElement, tags);

            ReplaceText(xmlElement, docx, tags);
        }

        /// <summary>
        /// 渲染Table
        /// </summary>
        /// <param name="table"></param>
        /// <param name="docx"></param>
        /// <param name="tags"></param>
        /// <exception cref="NotSupportedException"></exception>
        private static void GenerateTable(Table table, WordprocessingDocument docx, Dictionary<string, object> tags)
        {
            var trs = table.Descendants<TableRow>().ToArray(); // remember toarray or system will loop OOM;

            foreach (var tr in trs)
            {
                var innerText = tr.InnerText.Replace("{{foreach", "").Replace("endforeach}}", "")
                    .Replace("{{if(", "").Replace(")if", "").Replace("endif}}", "");

                // 匹配list数据，格式“Items.PropName”
                var matchs = (Regex.Matches(innerText, "(?<={{).*?\\..*?(?=}})")
                    .Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value)).ToArray();
                if (matchs.Length > 0)
                {
                    //var listKeys = matchs.Select(s => s.Split('.')[0]).Distinct().ToArray();
                    //// TODO:
                    //// not support > 2 list in same tr
                    //if (listKeys.Length > 2)
                    //    throw new NotSupportedException("MiniWord doesn't support more than 2 list in same row");
                    //var listKey = listKeys[0];

                    var listLevelKeys = matchs.Select(s => s.Substring(0, s.LastIndexOf('.'))).Distinct().ToArray();
                    // TODO:
                    // not support > 2 list in same tr
                    if (listLevelKeys.Length > 2)
                        throw new NotSupportedException("MiniWord doesn't support more than 2 list in same row");

                    var tagObj = GetObjVal(tags, listLevelKeys[0]);

                    if(tagObj == null) continue;

                    if (tagObj is IEnumerable)
                    {
                        var attributeKey = matchs[0].Split('.')[0];
                        var list = tagObj as IEnumerable;

                        foreach (var item in list)
                        {
                            var dic = new Dictionary<string, object>(); //TODO: optimize


                            var newTr = tr.CloneNode(true);
                            if (item is IDictionary)
                            {
                                var es = (Dictionary<string, object>)item;
                                foreach (var e in es)
                                {
                                    var dicKey = $"{listLevelKeys[0]}.{e.Key}";
                                    dic[dicKey] = e.Value;
                                }
                            }
                            // 支持Obj.A.B.C...
                            else
                            {
                                var props = item.GetType().GetProperties();
                                foreach (var p in props)
                                {
                                    var dicKey = $"{listLevelKeys[0]}.{p.Name}";
                                    dic[dicKey] = p.GetValue(item);
                                }
                            }

                            ReplaceIfStatements(newTr, tags: dic);

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
                    else
                    {
                        var dic = new Dictionary<string, object>(); //TODO: optimize

                        var props = tagObj.GetType().GetProperties();
                        foreach (var p in props)
                        {
                            var dicKey = $"{listLevelKeys[0]}.{p.Name}";
                            dic[dicKey] = p.GetValue(tagObj);
                        }

                        ReplaceIfStatements(tr, tags: tagObj.ToDictionary());

                        ReplaceText(tr, docx, tags: dic);
                    }
                }
                else
                {
                    var matchTxtProp = new Regex(@"(?<={{).*?\.?.*?(?=}})").Match(innerText);
                    if(!matchTxtProp.Success) continue;

                    ReplaceText(tr, docx, tags);
                }
            }
        }


        /// <summary>
        /// 获取Obj对象指定的值
        /// </summary>
        /// <param name="objSource">数据源</param>
        /// <param name="propNames">属性名，如“A.B”</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static object GetObjVal(object objSource, string propNames)
        {
            return GetObjVal(objSource, propNames.Split('.'));
        }

        /// <summary>
        /// 获取Obj对象指定的值
        /// </summary>
        /// <param name="objSource">数据源</param>
        /// <param name="propNames">属性名，如“A.B”即[0]A,[1]B</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static object GetObjVal(object objSource, string[] propNames)
        {
            var nextPropNames = propNames.Skip(1).ToArray();
            if (objSource is IDictionary)
            {
                var dict = (IDictionary)objSource;
                if (dict.Contains(propNames[0]))
                {
                    var val = dict[propNames[0]];
                    if(propNames.Length >1)
                        return GetObjVal(dict[propNames[0]], nextPropNames);
                    else return val;
                }
                return null;
            }
            // todo objSource = list
            var prop1 = objSource.GetType().GetProperty(propNames[0]);
            if (prop1 == null)
                return null;

            var prop1Val = prop1.GetValue(objSource);
            // 如果propNames只有一级，则直接返回对应的值
            if (propNames.Length == 1)
                return prop1Val;
            return GetObjVal(prop1Val, nextPropNames);
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

                    var foreachTagContains = s.Split(new[] { foreachTag }, StringSplitOptions.None).Length - 1 ==
                                             s.Split(new[] { endForeachTag }, StringSplitOptions.None).Length - 1;
                    var ifTagContains = s.Split(new[] { ifTag }, StringSplitOptions.None).Length - 1 ==
                                        s.Split(new[] { endifTag }, StringSplitOptions.None).Length - 1;
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
                case double dtg when double.TryParse(value, out var doubleNumber):
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

                    break;
                case int itg when int.TryParse(value, out var intNumber):
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

                    break;
                case DateTime dttg when DateTime.TryParse(value, out var date):
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

        /// <summary>
        /// 替换单个paragraph属性值
        /// </summary>
        /// <param name="p"></param>
        /// <param name="docx"></param>
        /// <param name="tags"></param>
        private static void ReplaceText(Paragraph p, WordprocessingDocument docx, Dictionary<string, object> tags)
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
                        // 完全匹配
                        var isFullMatch = t.Text.Contains($"{{{{{tag.Key}}}}}");
                        // 层级匹配，如{{A.B.C.D}}
                        var partMatch = new Regex($".*{{{{({tag.Key}(\\.\\w+)+)}}}}.*").Match(t.Text);

                        if (!isFullMatch && tag.Value is List<MiniWordForeach> forTags)
                        {
                            if (forTags.Any(forTag => forTag.Value.Keys.Any(dictKey =>
                            {
                                var innerTag = "{{" + tag.Key + "." + dictKey + "}}";
                                return t.Text.Contains(innerTag);
                            })))
                            {
                                isFullMatch = true;
                            }
                        }

                        if (isFullMatch || partMatch.Success)
                        {
                            var key = isFullMatch ? tag.Key : partMatch.Groups[1].Value;
                            var value = isFullMatch ? tag.Value : GetObjVal(tags, key);

                            if (value is string[] || value is IList<string> || value is List<string>)
                            {
                                var vs = value as IEnumerable;
                                var currentT = t;
                                var isFirst = true;
                                foreach (var v in vs)
                                {
                                    var newT = t.CloneNode(true) as Text;
                                    // todo
                                    newT.Text = t.Text.Replace($"{{{{{key}}}}}", v?.ToString());
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
                            // todo 未验证嵌套对象的渲染
                            else if (value is List<MiniWordForeach> vs)
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
                                        // todo tag,Key
                                        newT.Text = newT.Text.Replace("{{" + tag.Key + "." + vv.Key + "}}", vv.Value.ToString());
                                    }

                                    newT.Text = EvaluateIfStatement(newT.Text);

                                    if (!string.IsNullOrEmpty(newT.Text))
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
                            else if (IsHyperLink(value))
                            {
                                AddHyperLink(docx, run, value);
                                t.Remove();
                            }
                            else if (value is MiniWordColorText || value is MiniWordColorText[])
                            {
                                if (value is MiniWordColorText)
                                {
                                    AddColorText(run, new[] { (MiniWordColorText)value });
                                }
                                else
                                {
                                    AddColorText(run, (MiniWordColorText[])value);
                                }
                                t.Remove();
                            }
                            else if (value is MiniWordPicture)
                            {
                                var pic = (MiniWordPicture)value;
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
                                var newText = value is DateTime ? ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss") : value?.ToString();
                                t.Text = t.Text.Replace($"{{{{{key}}}}}", newText);
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

        private static void ReplaceText(OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
        {
            var paragraphs = xmlElement.Descendants<Paragraph>().ToArray();
            foreach (var p in paragraphs)
            {
                ReplaceText(p,docx,tags);
            }
        }

        /// <summary>
        /// @foreach元素复制及填充
        /// </summary>
        /// <param name="xmlElement"></param>
        private static void ReplaceForeachStatements(OpenXmlElement xmlElement,WordprocessingDocument docx,Dictionary<string,object> data)
        {
            // 1. 先获取Foreach的元素
            var beginKey = "@foreach";
            var endKey = "@endforeach";

            var betweenEles = GetBetweenElements(xmlElement, beginKey, endKey, false);
            while (betweenEles?.Any() == true)
            {
                var beginParagraph =
                    xmlElement.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains(beginKey));
                var endParagraph =
                    xmlElement.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains(endKey));
                // 获取需循环的数据key
                var match = new Regex(@".*{{(\w+(\.\w+)*)}}.*").Match(beginParagraph.InnerText);
                if (!match.Success) throw new Exception($"@Foreach循环未找到对应数据");
                var foreachDataKey = match.Groups[1].Value;

                // 删除关键字文本行
                beginParagraph?.Remove();
                endParagraph?.Remove();
                // 循环体最后一个元素，用于新元素插入定位
                var lastEleInLoop = betweenEles.LastOrDefault();
                var copyLoopEles = betweenEles.Select(e => e.CloneNode(true)).ToList();
                // 需要循环的数据
                var foreachList = GetObjVal(data, foreachDataKey);
                if (foreachList is IList list)
                {
                    var loopEles = new List<OpenXmlElement>();
                    for (var i = 0; i < list.Count; i++)
                    {
                        var item = list[i];
                        var foreachDataDict = item.ToDictionary();
                        // 2. 渲染替换属性值{{}}，插入循环元素，再替换……
                        // 2.1 替换属性值
                        if (i == 0)
                            loopEles = new List<OpenXmlElement>(betweenEles);
                        foreach (var ele in loopEles)
                        {
                            if (ele is Table table)
                                GenerateTable(table, docx, foreachDataDict);
                            else if (ele is Paragraph p)
                            {
                                ReplaceText(p, docx, foreachDataDict);
                            }
                        }
                        // @if代码块替换
                        ReplaceIfStatements(xmlElement, loopEles, foreachDataDict);

                        // 2.2 新增一个循环体元素
                        if (list.Count - 1 > i)
                        {
                            loopEles.Clear();
                            foreach (var ele in copyLoopEles)
                            {
                                var newEle = ele.CloneNode(true);
                                xmlElement.InsertAfter(newEle, lastEleInLoop);
                                lastEleInLoop = newEle;
                                loopEles.Add(newEle);
                            }
                        }
                    }
                }

                betweenEles = GetBetweenElements(xmlElement, beginKey, endKey, false);
            }
            
            
        }
        
        /// <summary>
        /// 获取关键词之间的元素
        /// </summary>
        /// <param name="sourceElement"></param>
        /// <param name="beginKey"></param>
        /// <param name="endKey"></param>
        /// <param name="isCopyEle">是否克隆元素对象。true：返回克隆元素；false：返回原始元素</param>
        /// <returns></returns>
        private static List<OpenXmlElement> GetBetweenElements(OpenXmlElement sourceElement,string beginKey,string endKey,bool isCopyEle = true)
        {
            var beginParagraph =
                sourceElement.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains(beginKey));
            var beginIndex = sourceElement.Elements().ToList().IndexOf(beginParagraph);
            if (beginIndex < 0) return null;

            var result = new List<OpenXmlElement>();
            foreach (var element in sourceElement.Elements().Skip(beginIndex + 1))
            {
                result.Add(isCopyEle ? element.CloneNode(true) : element);
                if (element is Paragraph p && p.InnerText.Contains(endKey))
                {
                    // 移除endKey的paragraph
                    result.RemoveAt(result.Count - 1);
                    return result;
                }
            }
            return result;
        }

        /// <summary>
        /// @if处理逻辑
        /// </summary>
        /// <param name="rootXmlElement">根元素</param>
        /// <param name="elementList">包含@if-@end的元素集合</param>
        /// <param name="tags"></param>
        private static void ReplaceIfStatements(OpenXmlElement rootXmlElement, List<OpenXmlElement> elementList, Dictionary<string, object> tags)
        {
            var paragraphs = elementList.Where(e=>e is Paragraph).ToList();
            while (paragraphs.Any(s => s.InnerText.Contains("@if")))
            {
                var ifP = paragraphs.First( s => s.InnerText.Contains("@if"));
                var endIfP = paragraphs.First( s => s.InnerText.Contains("@endif"));

                var statement = ifP.InnerText.Split(' ');

                //var tagValue = tags[statement[1]] ?? "NULL";
                var tagValue1 = GetObjVal(tags, statement[1]) ?? "NULL";
                var tagValue2 = GetObjVal(tags, statement[3]) ?? statement[3];

                var checkStatement = statement.Length == 4 ? EvaluateStatement(tagValue1.ToString(), statement[2], tagValue2.ToString()) : !bool.Parse(tagValue1.ToString());

                if (!checkStatement)
                {
                    var paragraphIfIndex = elementList.FindIndex(a => a == ifP);
                    var paragraphEndIfIndex = elementList.FindIndex(a => a == endIfP);

                    for (int i = paragraphIfIndex + 1; i <= paragraphEndIfIndex - 1; i++)
                    {
                        if(rootXmlElement.ChildElements.Any(c=>c == elementList[i])) rootXmlElement.RemoveChild(elementList[i]);
                    }
                }
                if(rootXmlElement.ChildElements.Any(c => c == ifP))
                    rootXmlElement.RemoveChild(ifP);
                if (rootXmlElement.ChildElements.Any(c => c == endIfP))
                    rootXmlElement.RemoveChild(endIfP);
                paragraphs.Remove(ifP);
                paragraphs.Remove(endIfP);
            }
        }

        /// <summary>
        /// @if处理逻辑
        /// </summary>
        /// <param name="xmlElement">@if-endif的父元素</param>
        /// <param name="tags"></param>
        private static void ReplaceIfStatements(OpenXmlElement xmlElement, Dictionary<string, object> tags)
        {
            var descendants = xmlElement.Descendants().ToList();

            ReplaceIfStatements(xmlElement,descendants, tags);
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
        private static void AddColorText(Run run, MiniWordColorText[] miniWordColorTextArray)
        {
            RunProperties runPro = null;
            foreach (var miniWordColorText in miniWordColorTextArray)
            {
                runPro = new RunProperties();
                Text text = new Text(miniWordColorText.Text);
                Color color = new Color() { Val = miniWordColorText.FontColor?.Replace("#", "") };
                Shading shading = new Shading() { Fill = miniWordColorText.HighlightColor?.Replace("#", "") };
                runPro.Append(shading);
                runPro.Append(color);
                run.Append(runPro);
                run.Append(text);
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
    }
}
