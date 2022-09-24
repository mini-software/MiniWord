<div align="center">
<p><a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/v/MiniWord.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/dt/MiniWord.svg" alt=""></a>  
<a href="https://github.com/mini-software/MiniWord" rel="nofollow"><img src="https://img.shields.io/github/stars/mini-software/MiniWord?logo=github" alt="GitHub stars"></a> 
<a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version"></a>
</p>
</div>

---

<div align="center">
<p><strong><a href="README.md">English</a> | <a href="README.zh-CN.md">简体中文</a> | <a href="README.zh-Hant.md">繁體中文</a></strong></p>
</div>

---

<div align="center">
 您的 <a href="https://github.com/mini-software/MiniWord">Star</a> 和 <a href="https://miniexcel.github.io">赞助</a> 可以让 MiniWord 走更远
</div>

---

## QQ群(1群) : [813100564](https://qm.qq.com/cgi-bin/qm/qr?k=3OkxuL14sXhJsUimWK8wx_Hf28Wl49QE&jump_from=webapi) / QQ群(2群) : [579033769](https://jq.qq.com/?_wv=1027&k=UxTdB8pR)

----

## 介绍

MiniWord .NET Word模板引擎，藉由Word模板和数据简单、快速生成文件。

![image](https://user-images.githubusercontent.com/12729184/190835307-6cd87982-b5f3-4a79-9682-bdd1cc02a4ea.png)



## Getting Started

### 安装

- nuget link : https://www.nuget.org/packages/MiniWord
- Packge xml `<PackageReference Include="MiniWord" Version="0.4.0" />`
- Or .NET CLI : `dotnet add package MiniWord --version 0.4.0`

### 快速入门

模板遵循“所见即所得”的设计，模板和标签的样式会被完全保留

```csharp
var value = new Dictionary<string, object>(){["title"] = "Hello MiniWord"};
MiniSoftware.MiniWord.SaveAsByTemplate(outputPath, templatePath, value);
```

![image](https://user-images.githubusercontent.com/12729184/190875707-6c5639ab-9518-4dc1-85d8-81e20af465e8.png)

### 输入、输出

- 输入系统支持模版路径或是Byte[]
- 输出支持文件路径、Byte[]、Stream

```csharp
SaveAsByTemplate(string path, string templatePath, Dictionary<string, object> value)
SaveAsByTemplate(string path, byte[] templateBytes, Dictionary<string, object> value)
SaveAsByTemplate(this Stream stream, string templatePath, Dictionary<string, object> value)
SaveAsByTemplate(this Stream stream, byte[] templateBytes, Dictionary<string, object> value)
```



## 标签

MiniWord 使用类似 Vue, React 的模版字串 `{{tag}}`，只需要确保 tag 与 value 参数的 key 一样`(大小写敏感)`，系统会自动替换字串。

### 文本

```csharp
{{tag}}
```



##### 代码例子

```csharp
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
```

##### 模版

![image](https://user-images.githubusercontent.com/12729184/190834360-39b4b799-d523-4b7e-9331-047a61fd5eb9.png)

##### 导出

![image](https://user-images.githubusercontent.com/12729184/190834455-ba065211-0f9d-41d1-9b7a-5d9e96ac2eff.png)

### 图片

标签值为 `MiniWordPicture` 类别

##### 代码例子

```csharp
var value = new Dictionary<string, object>()
{
    ["Logo"] = new MiniWordPicture() { Path= PathHelper.GetFile("DemoLogo.png"), Width= 180, Height= 180 }
};
MiniWord.SaveAsByTemplate(path, templatePath, value);
```



##### 模版

![image](https://user-images.githubusercontent.com/12729184/190647953-6f9da393-e666-4658-a56d-b3a7f13c0ea1.png)

##### 导出

![image](https://user-images.githubusercontent.com/12729184/190648179-30258d82-723d-4266-b711-43f132d1842d.png)

### 列表

标签值为 `string[]` 或是 `IList<string>`类别

##### 代码例子

```csharp
var value = new Dictionary<string, object>()
{
    ["managers"] = new[] { "Jack" ,"Alan"},
    ["employees"] = new[] { "Mike" ,"Henry"},
};
MiniWord.SaveAsByTemplate(path, templatePath, value);
```

##### 模版

![image](https://user-images.githubusercontent.com/12729184/190645513-230c54f3-d38f-47af-b844-0c8c1eff2f52.png)

##### 导出

![image](https://user-images.githubusercontent.com/12729184/190645704-1f6405e9-71e3-45b9-aa99-2ba52e5e1519.png)

### 表格

标签值为 `IEmerable<Dictionary<string,object>>`类别

##### 代码例子

```csharp
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
```

##### 模版

![image](https://user-images.githubusercontent.com/12729184/190843632-05bb6459-f1c1-4bdc-a79b-54889afdfeea.png)


##### 导出

![image](https://user-images.githubusercontent.com/12729184/190843663-c00baf16-21f2-4579-9d08-996a2c8c549b.png)



## 其他

### POCO or dynamic 参数

v0.5.0 支持 POCO 或 dynamic parameter

```csharp
var value = new { title = "Hello MiniWord" };
MiniWord.SaveAsByTemplate(outputPath, templatePath, value);
```


### HyperLink

我们可以尝试使用 `MiniWodrHyperLink` 类，用模板测试替换为超链接。

`MiniWordHyperLink` 提供了两个主要参数。

* Url： HyperLink URI 目标路径
* 文字：超链接文字

```csharp
var value = new 
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
```

## 例子



#### ASP.NET Core 3.1 API Export

```cs
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using MiniSoftware;

public class Program
{
    public static void Main(string[] args) => CreateHostBuilder(args).Build().Run();

    public static IHostBuilder CreateHostBuilder(string[] args) => Host.CreateDefaultBuilder(args).ConfigureWebHostDefaults(webBuilder => webBuilder.UseStartup<Startup>());
}

public class Startup
{
    public void ConfigureServices(IServiceCollection services) => services.AddMvc();
    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        app.UseStaticFiles();
        app.UseRouting();
        app.UseEndpoints(endpoints =>
        {
            endpoints.MapControllerRoute(
                name: "default",
                pattern: "{controller=api}/{action=Index}/{id?}");
        });
    }
}

public class ApiController : Controller
{
    public IActionResult Index()
    {
        return new ContentResult
        {
            ContentType = "text/html",
            StatusCode = (int)HttpStatusCode.OK,
            Content = @"<html><body>
<a href='api/DownloadWordFromTemplatePath'>DownloadWordFromTemplatePath</a><br>
<a href='api/DownloadWordFromTemplateBytes'>DownloadWordFromTemplateBytes</a><br>
</body></html>"
        };
    }

    static Dictionary<string, object> defaultValue = new Dictionary<string, object>()
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

    public IActionResult DownloadWordFromTemplatePath()
    {
        string templatePath = "TestTemplateComplex.docx";

        Dictionary<string, object> value = defaultValue;

        MemoryStream memoryStream = new MemoryStream();
        MiniWord.SaveAsByTemplate(memoryStream, templatePath, value);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        {
            FileDownloadName = "demo.docx"
        };
    }

    private static Dictionary<string, Byte[]> TemplateBytesCache = new Dictionary<string, byte[]>();

    static ApiController()
    {
        string templatePath = "TestTemplateComplex.docx";
        byte[] bytes = System.IO.File.ReadAllBytes(templatePath);
        TemplateBytesCache.Add(templatePath, bytes);
    }

    public IActionResult DownloadWordFromTemplateBytes()
    {
        byte[] bytes = TemplateBytesCache["TestTemplateComplex.docx"];

        Dictionary<string, object> value = defaultValue;

        MemoryStream memoryStream = new MemoryStream();
        MiniWord.SaveAsByTemplate(memoryStream, bytes, value);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        {
            FileDownloadName = "demo.docx"
        };
    }
}
```







## 支持我 : [Donate Link](https://miniexcel.github.io/)

<a href="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg"><img src="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg" alt="wechat" width="200px" height="300px">
</a> 
<a href="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg">
    <img src="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg" alt="alipay" width="200px" height="300px"></a>

## 常见问题

### 模版字串没有生效

建议 `{{tag}}` 复制重新整串复制贴上，有时打字 word 在底层 `{{}}`会被切开变成`<w:t>{</w:t><w:t>{<w:/t><w:t>Tag</w:t><w:t>}</w:t><w:t>}<w:/t>` 如图片

![image](https://user-images.githubusercontent.com/12729184/190683025-fbf1bfa3-a34a-4af9-a8d3-30c6807d229c.png)



