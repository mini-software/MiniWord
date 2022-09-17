using MiniSoftware;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace MiniWordTests
{
    public class IssueTests
    {

        /// <summary>
        /// [Support table generate · Issue #13 · mini-software/MiniWord]
        /// (https://github.com/mini-software/MiniWord/issues/13)
        /// </summary>
        [Fact]
        public void TestIssue13()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestExpenseDemo.docx");
            var value = new Dictionary<string, object>()
            {
                ["TripHs"] = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        { "sDate",DateTime.Parse("2022-09-08 08:30:00")},
                        { "eDate",DateTime.Parse("2022-09-08 15:00:00")},
                        { "How","Discussion requirement part1"},
                        { "Photo",new MiniWordPicture() { Path = PathHelper.GetFile("DemoExpenseMeeting02.png"), Width = 160, Height = 90 }},
                    },
                    new Dictionary<string, object>
                    {
                        { "sDate",DateTime.Parse("2022-09-09 08:30:00")},
                        { "eDate",DateTime.Parse("2022-09-09 17:00:00")},
                        { "How","Discussion requirement part2 and development"},
                        { "Photo",new MiniWordPicture() { Path = PathHelper.GetFile("DemoExpenseMeeting01.png"), Width = 160, Height = 90 }},
                    },
                }
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            //System.Diagnostics.Process.Start("explorer.exe", path);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains(@"Discussion requirement part2 and development", xml);
            Assert.Contains(@"Discussion requirement part1", xml);
        }

        [Fact]
        public void TestDemo01_Tag_Text()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestExpenseDemo.docx");
            var value = new Dictionary<string, object>()
            {
                ["Name"] = "Jack",
                ["Department"] = "IT Department",
                ["Purpose"] = "Shanghai site needs a new system to control HR system.",
                ["StartDate"] = DateTime.Parse("2022-09-07 08:30:00"),
                ["EndDate"] = DateTime.Parse("2022-09-15 15:30:00"),
                ["Approved"] = true,
                ["Total_Amount"] = 123456,
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
        }

        /// <summary>
        /// [System.InvalidOperationException: 'The parent of this element is null.' · Issue #12 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/12)
        /// </summary>
        [Fact]
        public void TestIssue12()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            var value = new Dictionary<string, object>()
            {
                ["Company_Name"] = "MiniSofteware\n",
                ["Name"] = "Jack",
                ["CreateDate"] = new DateTime(2021, 01, 01),
                ["VIP"] = true,
                ["Points"] = 123,
                ["APP"] = "Demo APP\n",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains(@"<w:t>MiniSofteware", xml);
            Assert.Contains(@"<w:br />", xml);
        }

        [Fact]
        public void TestIssueDemo03()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestDemo02.docx");
            var value = new Dictionary<string, object>()
            {
                ["FullName"] = "Julian Anderson",
                ["Title"] = "IT Manager",
                ["Phone"] = "+86 1234567890",
                ["Mail"] = "shps95100@gmail.com",
                ["Education"] = "Michigan State University | From Aug 2013 to May 2015",
                ["Major"] = "Computer Science",
                ["Favorites"] = "Music、Programing、Design",
                ["Skills"] = new[] { "- Photoshop", "- InDesign", "- MS Office", "- HTML 5", "- CSS 3" },
                ["Address"] = "1234, White Home, Road-12/ABC Street-13, New York, USA, 12345",
                ["AboutMe"] = "Hi, I’m Julian Anderson dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled the industry's standard dummy.",
                ["Birthday"] = "1993-09-26",
                ["Experiences"] = @"# SENIOR UI/UX DEVELOPER & DESIGNER
◼ The Matrix Media Limited | From May 2013 to May 2015
Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took.

◼ JUNIOR UI/UX DEVELOPER & DESIGNER
Linux OS Interface Limited | From Jan 2010 to Feb 2013
Lorem Ipsum has been the industry's standard dummy text 
ever since the 1500s, when an unknown printer took.

◼ TEAM LEADER & CORE GRAPHIC DESIGNER
Apple OS Interface Limited | From Jan 2008 to Feb 2010
Lorem Ipsum has been the industry's standard dummy text 
ever since the 1500s, when an unknown printer took.

◼ JUNIOR UI/UX DEVELOPER & DESIGNER
Apple OS Interface Limited | From Jan 2008 to Feb 2010
Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took.

◼ JUNIOR UI/UX DEVELOPER & DESIGNER
Apple OS Interface Limited | From Jan 2008 to Feb 2010
Lorem Ipsum has been the industry's standard dummy text 
ever since the 1500s, when an unknown printer took.
",
                ["Image"] = new MiniWordPicture() { Path = PathHelper.GetFile("demo01.png"), Width = 160, Height = 90 },
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            //System.Diagnostics.Process.Start("explorer.exe", path);
        }

        /// <summary>
        /// [support array list string to generate multiple row · Issue #11 · mini-software/MiniWord]
        /// (https://github.com/mini-software/MiniWord/issues/11)
        /// </summary>
        [Fact]
        public void TestIssue11()
        {
            {
                var path = PathHelper.GetTempFilePath();
                var templatePath = PathHelper.GetFile("TestIssue11.docx");
                var value = new Dictionary<string, object>()
                {
                    ["managers"] = new[] { "Jack", "Alan" },
                    ["employees"] = new[] { "Mike", "Henry" },
                };
                MiniWord.SaveAsByTemplate(path, templatePath, value);
                var xml = Helpers.GetZipFileContent(path, "word/document.xml");
                Assert.Contains("Jack", xml);
            }
            {
                var path = PathHelper.GetTempFilePath();
                var templatePath = PathHelper.GetFile("TestIssue11.docx");
                var value = new Dictionary<string, object>()
                {
                    ["managers"] = new List<string> { "Jack", "Alan" },
                    ["employees"] = new List<string> { "Mike", "Henry" },
                };
                MiniWord.SaveAsByTemplate(path, templatePath, value);
                var xml = Helpers.GetZipFileContent(path, "word/document.xml");
                Assert.Contains("Jack", xml);
            }
        }



        /// <summary>
        /// [Support image · Issue #3 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/3)
        /// </summary>
        [Fact]
        public void TestIssue3()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicImage.docx");
            var value = new Dictionary<string, object>()
            {
                ["Logo"] = new MiniWordPicture() { Path = PathHelper.GetFile("DemoLogo.png"), Width = 180, Height = 180 }
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains("<w:drawing>", xml);
        }

        /// <summary>
        /// [Fuzzy Regex replace similar key · Issue #5 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/5)
        /// </summary>
        [Fact]
        public void TestIssue5()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            var value = new Dictionary<string, object>()
            {
                ["Name"] = "Jack",
                ["Company_Name"] = "MiniSofteware",
                ["CreateDate"] = new DateTime(2021, 01, 01),
                ["VIP"] = true,
                ["Points"] = 123,
                ["APP"] = "Demo APP",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            //Console.WriteLine(path);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.DoesNotContain("Jack Demo APP Account Data", xml);
            Assert.Contains("MiniSofteware Demo APP Account Data", xml);
        }

        /// <summary>
        /// [Paragraph replace by tag · Issue #4 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/4)
        /// </summary>
        [Fact]
        public void TestIssue4()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            var value = new Dictionary<string, object>()
            {
                ["Company_Name"] = "MiniSofteware",
                ["Name"] = "Jack",
                ["CreateDate"] = new DateTime(2021, 01, 01),
                ["VIP"] = true,
                ["Points"] = 123,
                ["APP"] = "Demo APP",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
        }
    }

    internal static class Helpers
    {
        internal static string GetZipFileContent(string zipPath, string filePath)
        {
            var ns = new XmlNamespaceManager(new NameTable());
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            using (var stream = File.OpenRead(zipPath))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
            {
                var entry = archive.Entries.Single(w => w.FullName == filePath);
                using (var sheetStream = entry.Open())
                {
                    var doc = XDocument.Load(sheetStream);
                    return doc.ToString();
                }
            }
        }
    }

    internal static class PathHelper
    {
        public static string GetFile(string fileName, string folderName = "docx")
        {
            return $@"../../../../../samples/{folderName}/{fileName}";
        }

        public static string GetTempPath(string extension = "docx")
        {
            var method = (new System.Diagnostics.StackTrace()).GetFrame(1).GetMethod();

            var path = Path.Combine(Path.GetTempPath(), $"{method.DeclaringType.Name}_{method.Name}.{extension}").Replace("<", string.Empty).Replace(">", string.Empty);
            if (File.Exists(path))
                File.Delete(path);
            return path;
        }

        public static string GetTempFilePath(string extension = "docx")
        {
            return Path.GetTempPath() + Guid.NewGuid().ToString() + "." + extension;
        }
    }
}
