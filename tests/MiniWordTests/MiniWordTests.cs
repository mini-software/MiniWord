using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using MiniSoftware;
using Xunit;

namespace MiniWordTests
{
    public class MiniWordTests
    {
        [Fact]
        public void TestForeachLoopInTables()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestForeachInTablesDemo.docx");
            var value = new Dictionary<string, object>()
            {
                ["TripHs"] = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        { "sDate", DateTime.Parse("2022-09-08 08:30:00") },
                        { "eDate", DateTime.Parse("2022-09-08 15:00:00") },
                        { "How", "Discussion requirement part1" },
                        {
                            "Details", new List<MiniWordForeach>()
                            {
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Air"},
                                        {"Value", "Airplane"}
                                    },
                                    Separator = " | "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Parking"},
                                        {"Value", "Car"}
                                    },
                                    Separator = " / "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Hotel"},
                                        {"Value", "Room"}
                                    },
                                    Separator = ", "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Food"},
                                        {"Value", "Plate"}
                                    },
                                    Separator = ""
                                }
                            }
                        }
                    },
                    new Dictionary<string, object>
                    {
                        { "sDate", DateTime.Parse("2022-09-09 08:30:00") },
                        { "eDate", DateTime.Parse("2022-09-09 17:00:00") },
                        { "How", "Discussion requirement part2 and development" },
                        {
                            "Details", new List<MiniWordForeach>()
                            {
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Air"},
                                        {"Value", "Airplane"}
                                    },
                                    Separator = " | "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Parking"},
                                        {"Value", "Car"}
                                    },
                                    Separator = " / "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Hotel"},
                                        {"Value", "Room"}
                                    },
                                    Separator = ", "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Food"},
                                        {"Value", "Plate"}
                                    },
                                    Separator = ""
                                }
                            }
                        }
                    }
                }
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            //System.Diagnostics.Process.Start("explorer.exe", path);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains(@"Discussion requirement part2 and development", xml);
            Assert.Contains(@"Discussion requirement part1", xml);
            Assert.Contains(
                "Air way to the Airplane | Parking way to the Car / Hotel way to the Room, Food way to the Plate", xml);
        }
        
        [Fact]
        public void MiniWordIfStatement_FirstIf()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestIfStatement.docx");
            var value = new Dictionary<string, object>()
            {
                ["Name"] = new List<MiniWordHyperLink>(){
                    new MiniWordHyperLink(){
                        Url = "https://google.com",
                        Text = "測試連結22!!"
                    },
                    new MiniWordHyperLink(){
                        Url = "https://google1.com",
                        Text = "測試連結11!!"
                    }
                },
                ["Company_Name"] = "MiniSofteware",
                ["CreateDate"] = new DateTime(2021, 01, 01),
                ["VIP"] = true,
                ["Points"] = 123,
                ["APP"] = "Demo APP",
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            //Console.WriteLine(path);
            var docXml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains("First if chosen: MiniSofteware", docXml);
            Assert.DoesNotContain("Second if chosen: MaxiSoftware", docXml);
            Assert.Contains("Points are greater than 100", docXml);
            Assert.Contains("CreateDate is not less than 2021", docXml);
            Assert.DoesNotContain("CreateDate is not greater than 2021", docXml);
        }
        
        [Fact]
        public void TestForeachLoopInTablesWithIfStatement()
        {
            var path = PathHelper.GetTempFilePath();
            var templatePath = PathHelper.GetFile("TestForeachInTablesWithIfStatementDemo.docx");
            var value = new Dictionary<string, object>()
            {
                ["TripHs"] = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        { "sDate", DateTime.Parse("2022-09-08 08:30:00") },
                        { "eDate", DateTime.Parse("2022-09-08 15:00:00") },
                        { "How", "Discussion requirement part1" },
                        {
                            "Details", new List<MiniWordForeach>()
                            {
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Air"},
                                        {"Value", "Airplane"}
                                    },
                                    Separator = " | "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Parking"},
                                        {"Value", "Car"}
                                    },
                                    Separator = " / "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Hotel"},
                                        {"Value", "Room"}
                                    },
                                    Separator = ", "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Food"},
                                        {"Value", "Plate"}
                                    },
                                    Separator = ""
                                }
                            }
                        }
                    },
                    new Dictionary<string, object>
                    {
                        { "sDate", DateTime.Parse("2022-09-09 08:30:00") },
                        { "eDate", DateTime.Parse("2022-09-09 17:00:00") },
                        { "How", "Discussion requirement part2 and development" },
                        {
                            "Details", new List<MiniWordForeach>()
                            {
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Air"},
                                        {"Value", "Airplane"}
                                    },
                                    Separator = " | "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Parking"},
                                        {"Value", "Car"}
                                    },
                                    Separator = " / "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Hotel"},
                                        {"Value", "Room"}
                                    },
                                    Separator = ", "
                                },
                                new MiniWordForeach()
                                {
                                    Value = new Dictionary<string, object>()
                                    {
                                        {"Text", "Food"},
                                        {"Value", "Plate"}
                                    },
                                    Separator = ""
                                }
                            }
                        }
                    }
                }
            };
            MiniWord.SaveAsByTemplate(path, templatePath, value);
            //System.Diagnostics.Process.Start("explorer.exe", path);
            var xml = Helpers.GetZipFileContent(path, "word/document.xml");
            Assert.Contains(@"Discussion requirement part2 and development", xml);
            Assert.Contains(@"Discussion requirement part1", xml);
            Assert.Contains("Air way to the Airplane | Hotel way to the Room", xml);
        }
    }
}