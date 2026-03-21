using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using MiniSoftware;
using Xunit;

namespace MiniWordTests;

public class MiniWordTest4MeAsync
{
    [Fact]
    public async void TestForeachLoopInTablesAsync()
    {
        var path = "C:\\Users\\Administrator\\Downloads\\";// PathHelper.GetTempFilePath();
        var templatePath = "D:\\工作文档\\毅智科技\\导出模板\\信息标签模板A4_2.docx";//PathHelper.GetFile("TestForeachInTablesDemo.docx");


        var list = new List<Dictionary<string, object>>();

        foreach (var i in Enumerable.Range(1, 8))
        {
            var values = new Dictionary<string, object>()
            {
                ["u:name"] = "李哈哈" + i,
                ["u:sex"] = "男",
                ["r:userGroupParentName"] = "广州人才凯盛幼儿园",
                ["r:userGroupName"] = $"大{i}班",
                ["e:2613220640424581"] = i % 2 == 0 ? "男" : "女",
                ["e:2613186883781253"] = i % 2 == 0 ? "5" : "6",
                ["u:months"] = i,
                ["e:2660464484401477"] = i,
                ["s:2642862487273797"] = "维生素D",
                ["s:2642861958398277"] = "乙肝两对半",
                ["s:2642862864793925"] = "铁蛋白",

                ["qc1"] = "sdsdfds",
            };
            list.Add(values);
        }

        var value = new Dictionary<string, object>();
        value["list"] = list;


        await MiniWord.SaveAsByTemplateAsync(path + DateTime.Now.ToFileTimeUtc() + ".docx", templatePath, value);
        //System.Diagnostics.Process.Start("explorer.exe", path);
        //var xml = Helpers.GetZipFileContent(path + DateTime.Now.ToFileTimeUtc() + ".docx", "word/document.xml");
        //Assert.Contains(@"Discussion requirement part2 and development", xml);
        //Assert.Contains(@"Discussion requirement part1", xml);
        //Assert.Contains(
        //    "Air way to the Airplane | Parking way to the Car / Hotel way to the Room, Food way to the Plate", xml);
    }

    [Fact]
    public async void TestForeachLoopInTables2Async()
    {

        string path = "D:\\工作文档\\毅智科技\\导出模板\\3-6岁体检表2026_列表.docx";
        string outputPath = "C:\\Users\\Administrator\\Downloads\\";
        List<MemoryStream> streamList = new List<MemoryStream>();
        List<Dictionary<string, object>> valueList = new List<Dictionary<string, object>>();
        byte[] buffer = File.ReadAllBytes("C:\\Users\\Administrator\\Downloads\\qcode.png");

        foreach (var index in Enumerable.Range(1, 6))
        {
            var values = new Dictionary<string, object>()
            {
                ["u:name"] = "李哈哈_" + index,
                ["u:sex"] = "男",
                ["r:userGroupParentName"] = "广州人才凯盛幼儿园",
                ["r:userGroupName"] = "大一班",
                ["u:sex"] = "男",
                ["u:age"] = "5",
                ["u:months"] = "4",
                ["e:ti200020"] = "110",
                ["s:2642862487273797"] = "维生素D",
                ["s:ei100008"] = "乙肝两对半",
                ["s:ei100009"] = "铁蛋白",
                ["qc1"] = new MiniSoftware.MiniWordPicture(buffer, MiniSoftware.Common.Enums.Extension.Png, 50, 50)
            };
            //MemoryStream stream = new MemoryStream();
            //await MiniSoftware.MiniWord.SaveAsByTemplateAsync(stream, path, values);
            //streamList.Add(stream);
            valueList.Add(values);
        };
        Dictionary<string, object> listData = new Dictionary<string, object>()
        {
        };
        listData["list"] = valueList;
        MemoryStream stream = new MemoryStream();
        await MiniSoftware.MiniWord.SaveAsByTemplateAsync(outputPath + (DateTime.Now.ToFileTimeUtc()) + ".docx", path, listData);
    }

    [Fact]
    public async void MiniWordIfStatement_FirstIfAsync()
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
        await MiniWord.SaveAsByTemplateAsync(path, templatePath, value);
        //Console.WriteLine(path);
        var docXml = Helpers.GetZipFileContent(path, "word/document.xml");
        Assert.Contains("First if chosen: MiniSofteware", docXml);
        Assert.DoesNotContain("Second if chosen: MaxiSoftware", docXml);
        Assert.Contains("Points are greater than 100", docXml);
        Assert.Contains("CreateDate is not less than 2021", docXml);
        Assert.DoesNotContain("CreateDate is not greater than 2021", docXml);
    }
    
    [Fact]
    public async void TestForeachLoopInTablesWithIfStatementAsync()
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
        await MiniWord.SaveAsByTemplateAsync(path, templatePath, value);
        //System.Diagnostics.Process.Start("explorer.exe", path);
        var xml = Helpers.GetZipFileContent(path, "word/document.xml");
        Assert.Contains(@"Discussion requirement part2 and development", xml);
        Assert.Contains(@"Discussion requirement part1", xml);
        Assert.Contains("Air way to the Airplane | Hotel way to the Room", xml);
    }
}