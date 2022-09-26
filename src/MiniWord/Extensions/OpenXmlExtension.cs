namespace MiniSoftware.Extensions
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Dynamic;
    using System.Linq;
    using System.Text;

    internal static class OpenXmlExtension
    {
        /// <summary>
        /// 高级搜索：得到段落里面的连续字符串
        /// </summary>
        /// <param name="paragraph">段落</param>
        /// <returns>Item1：连续文本；Item2：块；Item3：块文本</returns>
        internal static List<Tuple<string, List<Run>, List<Text>>> GetContinuousString(this Paragraph paragraph)
        {
            List<Tuple<string, List<Run>, List<Text>>> tuples = new List<Tuple<string, List<Run>, List<Text>>>();
            if (paragraph == null)
                return tuples;

            var sb = new StringBuilder();
            var runs = new List<Run>();
            var texts = new List<Text>();

            //段落：所有子级
            foreach (var pChildElement in paragraph.ChildElements)
            {
                //块
                if (pChildElement is Run run)
                {
                    //文本块
                    if (run.IsText())
                    {
                        var text = run.GetFirstChild<Text>();
                        runs.Add(run);
                        texts.Add(text);
                        sb.Append(text.InnerText);
                    }
                    else
                    {
                        if (runs.Any())
                            tuples.Add(new Tuple<string, List<Run>, List<Text>>(sb.ToString(), runs, texts));

                        sb = new StringBuilder();
                        runs = new List<Run>();
                        texts = new List<Text>();
                    }
                }
                //公式，书签...
                else
                {
                    //跳过的类型
                    if (pChildElement is BookmarkStart || pChildElement is BookmarkEnd)
                    {

                    }
                    else
                    {
                        if (runs.Any())
                            tuples.Add(new Tuple<string, List<Run>, List<Text>>(sb.ToString(), runs, texts));

                        sb = new StringBuilder();
                        runs = new List<Run>();
                        texts = new List<Text>();
                    }
                }
            }

            if (runs.Any())
                tuples.Add(new Tuple<string, List<Run>, List<Text>>(sb.ToString(), runs, texts));

            sb = new StringBuilder();
            runs = new List<Run>();
            texts = new List<Text>();

            return tuples;
        }

        /// <summary>
        /// 整理字符串到连续字符串块中
        /// </summary>
        /// <param name="texts">连续字符串块</param>
        /// <param name="text">待整理字符串</param>
        internal static void TrimStringToInContinuousString(this IEnumerable<Text> texts, string text)
        {
            /*
            //假如块为：[A][BC][DE][FG][H]
            //假如替换：[AB][E][GH]
            //优化块为：[AB][C][DE][FGH][]
             */

            var index = string.Concat(texts.SelectMany(o => o.Text)).IndexOf(text);
            if (index > 0)
            {
                int i = -1;
                int addLengg = 0;
                bool isbr = false;
                foreach (var textWord in texts)
                {
                    if (addLengg > 0)
                    {
                        isbr = true;
                        var leng = textWord.Text.Length;

                        if (addLengg - leng > 0)
                        {
                            addLengg -= leng;
                            textWord.Text = "";
                        }
                        else if (addLengg - leng == 0)
                        {
                            textWord.Text = "";
                            break;
                        }
                        else
                        {
                            textWord.Text = textWord.Text.Substring(addLengg);
                        }
                    }
                    else if (isbr)
                    {
                        break;
                    }
                    else
                    {
                        i += textWord.Text.Length;
                        //开始包含
                        if (i >= index)
                        {
                            //全部包含
                            if (textWord.Text.Contains(text))
                            {
                                break;
                            }
                            //部分包含
                            else
                            {
                                var str1 = textWord.Text.Substring(0, i - index + 1);
                                if (i == index)
                                    str1 = "";

                                var str2 = str1 + text;

                                addLengg = str2.Length - textWord.Text.Length;
                                textWord.Text = str2;
                            }
                        }
                    }
                }
            }
        }


        internal static bool IsText(this Run run)
        {
            return run.Elements().All(o => o is Text || o is RunProperties);
        }
    }
}