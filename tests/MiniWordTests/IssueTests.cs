using MiniSoftware;
using System;
using System.Collections.Generic;
using Xunit;
using System.Dynamic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MiniWordTests
{
    public class IssueTests
    {

        /// <summary>
        /// [Text color multiple tags format error · Issue #37 · mini-software/MiniWord]
        /// (https://github.com/mini-software/MiniWord/issues/37)
        /// </summary>
        [Fact]
        public void TestIssue37()
        {
            {
                var path = PathHelper.GetTempFilePath();
                var templatePath = PathHelper.GetFile("TestIssue37.docx");
                var value = new Dictionary<string, object>
                {
                    ["Content"] = "Test",
                    ["Content2"] = "Test2",
                };
                MiniWord.SaveAsByTemplate(path, templatePath, value);
                var xml = Helpers.GetZipFileContent(path, "word/document.xml");
                Assert.Contains(@"Test", xml);
                Assert.Contains(@"Test2", xml);
            }

            {
                var path = PathHelper.GetTempFilePath();
                var templatePath = PathHelper.GetFile("TestIssue37.docx");
                var value = new Dictionary<string, object>
                {
                    ["Content"] = new MiniWordHyperLink()
                    {
                        Url = "https://google.com",
                        Text = "Test1!!"
                    },
                    ["Content2"] = new MiniWordHyperLink()
                    {
                        Url = "https://google.com",
                        Text = "Test2!!"
                    },
                };
                MiniWord.SaveAsByTemplate(path, templatePath, value);
                var xml = Helpers.GetZipFileContent(path, "word/document.xml");
                Assert.Contains(@"Test", xml);
                Assert.Contains(@"Test2", xml);
            }
        }

        [Fact]
        public void TestDemo04()
        {
            var outputPath = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestDemo04.docx");
            var value = new Dictionary<string, object>() { ["title"] = "Hello MiniWord" };
            MiniSoftware.MiniWord.SaveAsByTemplate(outputPath, templatePath, value);
        }

        [Fact]
        public void TestDemo04_new()
        {
            var outputPath = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestDemo04.docx");
            var value = new { title = "Hello MiniWord" };
            MiniWord.SaveAsByTemplate(outputPath, templatePath, value);
        }


        [Fact]
        public void TestIssue18()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestIssue18.docx");

            {
                var value = new Dictionary<string, object>()
                {
                    ["title"] = "FooCompany",
                    ["managers"] = new List<Dictionary<string, object>> {
                        new Dictionary<string, object>{{"name","Jack"},{ "department", "HR" } },
                        new Dictionary<string, object> {{ "name", "Loan"},{ "department", "IT" } }
                    },
                    ["employees"] = new List<Dictionary<string, object>> {
                        new Dictionary<string, object>{{ "name", "Wade" },{ "department", "HR" } },
                        new Dictionary<string, object> {{ "name", "Felix" },{ "department", "HR" } },
                        new Dictionary<string, object>{{ "name", "Eric" },{ "department", "IT" } },
                        new Dictionary<string, object> {{ "name", "Keaton" },{ "department", "IT" } }
                    }
                };
                MiniWord.SaveAsByTemplate(path, templatePath, value);
                var xml = Helpers.GetZipFileContent(path, "word/document.xml");
                Assert.Contains(@"<w:t>Keaton", xml);
                Assert.Contains(@"<w:t>Eric", xml);
            }

            //Strong type
            {
                var value = new
                {
                    title = "FooCompany",
                    managers = new[]
                    {
                        new {name="Jack",department="HR" },
                        new {name="Loan",department="IT" },
                    },
                    employees = new[]
                    {
                        new {name="Jack",department="HR" },
                        new {name="Loan",department="HR" },
                        new {name="Eric",department="IT" },
                        new {name="Keaton",department="IT" },
                    },
                };
                MiniWord.SaveAsByTemplate(path, templatePath, value);
                var xml = Helpers.GetZipFileContent(path, "word/document.xml");
                Assert.Contains(@"<w:t>Keaton", xml);
                Assert.Contains(@"<w:t>Eric", xml);
            }

            //Strong type
            {
                Foo value = new Foo()
                {
                    title = "FooCompany",
                    managers = new List<User>()
                    {
                        new User (){ name="Jack",department="HR"},
                        new User (){ name="Loan",department="IT"},
                    },
                        employees = new List<User>()
                    {
                        new User (){ name="Jack",department="HR"},
                        new User (){ name="Loan",department="HR"},
                        new User (){ name="Eric",department="IT"},
                        new User (){ name="Keaton",department="IT"},
                    },
                };
                MiniWord.SaveAsByTemplate(path, templatePath, value);
                var xml = Helpers.GetZipFileContent(path, "word/document.xml");
                Assert.Contains(@"<w:t>Keaton", xml);
                Assert.Contains(@"<w:t>Eric", xml);
            }
        }

        /// <summary>
        /// [Split template string like `<w:t>{</w:t><w:t>{<w:/t><w:t>Tag</w:t><w:t>}</w:t><w:t>}<w:/t>` problem · Issue #17 · mini-software/MiniWord]
        /// (https://github.com/mini-software/MiniWord/issues/17)
        /// </summary>
        [Fact]
        public void TestIssue17()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestIssue17.docx");
            var value = new Dictionary<string, object>()
            {
                ["Content"] = "Test",
                ["Content2"] = "Test2",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains(@"<w:t>Test", xml);
            Assert.Contains(@"<w:t>Test2", xml);
        }

        /// <summary>
        /// [Split template string like `<w:t>{</w:t><w:t>{<w:/t><w:t>Tag</w:t><w:t>}</w:t><w:t>}<w:/t>` problem · Issue #17 · mini-software/MiniWord]
        /// (https://github.com/mini-software/MiniWord/issues/17)
        /// </summary>
        [Fact]
        public void TestIssue17_new()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestIssue17.docx");
            var value = new 
            {
                Content = "Test",
                Content2 = "Test2",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains(@"<w:t>Test", xml);
            Assert.Contains(@"<w:t>Test2", xml);
        }


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

        /// <summary>
        /// [Support table generate · Issue #13 · mini-software/MiniWord]
        /// (https://github.com/mini-software/MiniWord/issues/13)
        /// </summary>
        [Fact]
        public void TestIssue13_new()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestExpenseDemo.docx");
            var value = new 
            {
                TripHs = new List<Dictionary<string, object>>
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


        [Fact]
        public void TestDemo01_Tag_Text_new()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestExpenseDemo.docx");
            var value = new
            {
                Name = "Jack",
                Department = "IT Department",
                Purpose = "Shanghai site needs a new system to control HR system.",
                StartDate = DateTime.Parse("2022-09-07 08:30:00"),
                EndDate = DateTime.Parse("2022-09-15 15:30:00"),
                Approved = true,
                Total_Amount = 123456,
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

        /// <summary>
        /// [System.InvalidOperationException: 'The parent of this element is null.' · Issue #12 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/12)
        /// </summary>
        [Fact]
        public void TestIssue12_dynamic()
        {
            
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            dynamic value = new ExpandoObject();
            value.Company_Name = "MiniSofteware\n";
            value.Name = "Jack";
            value.CreateDate = new DateTime(2021, 01, 01);
            value.VIP = true;
            value.Points = 123;
            value.APP = "Demo APP\n";
            
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains(@"<w:t>MiniSofteware", xml);
            Assert.Contains(@"<w:br />", xml);
        }

        /// <summary>
        /// [System.InvalidOperationException: 'The parent of this element is null.' · Issue #12 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/12)
        /// </summary>
        [Fact]
        public void TestIssue12_new()
        {

            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            object value = new {
                Company_Name = "MiniSofteware\n",
                Name = "Jack",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123,
                APP = "Demo APP\n",
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

        [Fact]
        public void TestIssueDemo03_dynamic()
        {
            
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestDemo02.docx");
            dynamic value = new ExpandoObject();
            value.FullName = "Julian Anderson";
            value.Title = "IT Manager";
            value.Phone = "+86 1234567890";
            value.Mail = "shps95100@gmail.com";
            value.Education = "Michigan State University | From Aug 2013 to May 2015";
            value.Major = "Computer Science";
            value.Favorites = "Music、Programing、Design";
            value.Skills = new[] { "- Photoshop", "- InDesign", "- MS Office", "- HTML 5", "- CSS 3" };
            value.Address = "1234, White Home, Road-12/ABC Street-13, New York, USA, 12345";
            value.AboutMe = "Hi, I’m Julian Anderson dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled the industry's standard dummy.";
            value.Birthday = "1993-09-26";
            value.Experiences = @"# SENIOR UI/UX DEVELOPER & DESIGNER
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
    ";
            value.Image = new MiniWordPicture() { Path = PathHelper.GetFile("demo01.png"), Width = 160, Height = 90 };
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
        /// [support array list string to generate multiple row · Issue #11 · mini-software/MiniWord]
        /// (https://github.com/mini-software/MiniWord/issues/11)
        /// </summary>
        [Fact]
        public void TestIssue11_new()
        {
            //{
            //    var path = PathHelper.GetTempFilePath();
            //    var templatePath = PathHelper.GetFile("TestIssue11.docx");
            //    var value = new 
            //    {
            //        managers = new[] { "Jack", "Alan" },
            //        employees = new[] { "Mike", "Henry" },
            //    };
            //    MiniWord.SaveAsByTemplate(path, templatePath, value);
            //    var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            //    Assert.Contains("Jack", xml);
            //}
            {
                var path = PathHelper.GetTempFilePath();
                var templatePath = PathHelper.GetFile("TestIssue11.docx");
                var value = new
                {
                    managers = new List<string> { "Jack", "Alan" },
                    employees = new List<string> { "Mike", "Henry" },
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
                ["Name"] = new MiniWordHyperLink(){
                    Url = "https://google.com",
                    Text = "測試連結!!"
                },
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
            Assert.Contains("MiniSofteware Demo APP Account Data", xml);
            Assert.Contains("<w:hyperlink w:tgtFrame=\"_blank\"", xml);
        }

        /// <summary>
        /// [Fuzzy Regex replace similar key · Issue #5 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/5)
        /// </summary>
        [Fact]
        public void TestIssue5_new()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            var value = new 
            {
                Name =new MiniWordHyperLink(){
                    Url = "https://google.com",
                    Text = "測試連結!!"
                },
                Company_Name = "MiniSofteware",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123,
                APP = "Demo APP",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            //Console.WriteLine(path);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.DoesNotContain("Jack Demo APP Account Data", xml);
            Assert.Contains("MiniSofteware Demo APP Account Data", xml);
            Assert.Contains("<w:hyperlink w:tgtFrame=\"_blank\"", xml);
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

        /// <summary>
        /// [Paragraph replace by tag · Issue #4 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/4)
        /// </summary>
        [Fact]
        public void TestIssue4_new()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            var value = new 
            {
                Company_Name = "MiniSofteware",
                Name = "Jack",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123,
                APP = "Demo APP",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
        }
        [Fact]
        public void TestColor()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
            var value = new
            {
                Company_Name = new MiniWordColorText { Text = "MiniSofteware", FontColor = "#eb70AB" },
                Name = new MiniWordColorText { Text = "Jack", HighlightColor = "#eb70AB" },
                CreateDate = new MiniWordColorText { Text = new DateTime(2021, 01, 01).ToString()
                    , HighlightColor = "#eb70AB", FontColor = "#ffffff" },
                VIP = true,
                Points = 123,
                APP = "Demo APP",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
        }


        #region Model:TestIssue18.docx
        public class Foo
        {
            public string title { get; set; }
            public List<User> managers { get; set; }
            public List<User> employees { get; set; }
        }
        public class User
        {
            public string name { get; set; }
            public string department { get; set; }
        }
        #endregion
    }
}
