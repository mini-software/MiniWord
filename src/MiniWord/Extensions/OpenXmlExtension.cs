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
                        sb.Append(text.Text);
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

            sb = null;
            runs = null;
            texts = null;

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

            var allTxtx = string.Concat(texts.SelectMany(o => o.Text));
            var indexState = allTxtx.IndexOf(text);
            if (indexState == -1)
                return;

            int indexEnd = indexState + text.Length - 1;
            List<Tuple<int, char>> yl = new List<Tuple<int, char>>(allTxtx.Length);
            int iRun = 0;
            int iIndex = 0;
            int iRunOf = -1;
            foreach (var item in texts)
            {
                foreach (var item2 in item.Text)
                {
                    if (indexState <= iIndex && iIndex <= indexEnd)
                    {
                        if (iRunOf == -1)
                            iRunOf = iRun;

                        yl.Add(new Tuple<int, char>(iRunOf, item2));
                    }
                    else
                    {
                        yl.Add(new Tuple<int, char>(iRun, item2));
                    }

                    iIndex++;
                }
                iRun++;
            }

            int i = 0;
            foreach (var item in texts)
            {
                item.Text = string.Concat(yl.Where(o => o.Item1 == i).Select(o => o.Item2));
                i++;
            }

        }


        internal static bool IsText(this Run run)
        {
            return run.Elements().All(o => o is Text || o is RunProperties);
        }
    }
}