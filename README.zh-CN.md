<div align="center">
<p><a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/v/MiniWord.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/dt/MiniWord.svg" alt=""></a>
<a href="https://github.com/mini-software/MiniWord" rel="nofollow"><img src="https://img.shields.io/github/stars/mini-software/MiniWord?logo=github" alt="GitHub stars"></a>
<a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version"></a>
<a href="https://deepwiki.com/mini-software/MiniWord"><img src="https://deepwiki.com/badge.svg" alt="Ask DeepWiki"></a>
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

### 二级列表

Tag 是 `IEnumerable<MiniWordForeach>` 类别.  使用方式`{{foreach` 和 `endforeach}}`.

##### Example

```csharp
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
                    }
                }
            }
        }
    }
};
MiniWord.SaveAsByTemplate(path, templatePath, value);
```

##### Template

![before_foreach](https://user-images.githubusercontent.com/38832863/220123955-063c9345-3998-4fd7-982c-8d1e3b48bbf8.PNG)

##### Result

![after_foreach](https://user-images.githubusercontent.com/38832863/220123960-913a7140-2fa2-415e-bb3e-456e04167382.PNG)

### 条件判断

 `@if` 和 `@endif` tags .

##### Example

```csharp
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
```

##### Template

![before_if](https://user-images.githubusercontent.com/38832863/220125429-7dd6ce94-35c6-478e-8903-064f9cf9361a.PNG)

##### Result

![after_if](https://user-images.githubusercontent.com/38832863/220125435-72ea24b4-2412-45de-961a-ad4b2134417b.PNG)

### 循环

 `@foreach` 和 `@endforeach` tags .

##### Example

```csharp
var value = new
{
    LoopData = new List<object>()
    {
        new {
            Type="类型A",
            Items = new List<object>() {new {Name = "A-1"}, new {Name = "A-2"},}
        },
        new
        {
            Type="类型B",
            Items = new List<object>() {new {Name = "B-1"}, new {Name = "B-2"}, new {Name = "B-3"},}
        },
    }
};
MiniWord.SaveAsByTemplate(path, templatePath, value);
```

##### Template

![1](https://github.com/user-attachments/assets/5d32241d-3977-46e7-b3de-cae130e5a653)

##### Result

![2](https://github.com/user-attachments/assets/69daa15e-4864-483e-b132-d8e867b6d1d1)

### 多彩字体

##### 代码例子

```csharp
var value = new
{
    Company_Name = new MiniWordColorText { Text = "MiniSofteware", FontColor = "#eb70AB", },
    Name = new[] {
        new MiniWordColorText { Text = "Ja", HighlightColor = "#eb70AB" },
        new MiniWordColorText { Text = "ck", HighlightColor = "#a56abe" }
    },
    CreateDate = new MiniWordColorText
    {
        Text = new DateTime(2021, 01, 01).ToString(),
        HighlightColor = "#eb70AB",
        FontColor = "#ffffff",
    },
    VIP = true,
    Points = 123,
    APP = "Demo APP",
};
MiniWord.SaveAsByTemplate(path, templatePath, value);
```





## 其他

### POCO or dynamic 参数

v0.5.0 支持 POCO 或 dynamic parameter

```csharp
var value = new { title = "Hello MiniWord" };
MiniWord.SaveAsByTemplate(outputPath, templatePath, value);
```

### 字体FontColor和HighlightColor
```csharp
var value = new
{
    Company_Name = new MiniWordColorText { Text = "MiniSofteware", FontColor = "#eb70AB" },
    Name = new MiniWordColorText { Text = "Jack", HighlightColor = "#eb70AB" },
    CreateDate = new MiniWordColorText { Text = new DateTime(2021, 01, 01).ToString(), HighlightColor = "#eb70AB", FontColor = "#ffffff" },
    VIP = true,
    Points = 123,
    APP = "Demo APP",
};
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

### 浮动图像

可以通过MiniWordPicture扩展参数配置图片悬浮环绕在文字上或文字下
`MiniWordPicture` 扩展参数。
* WrappingType： MiniWordPictureWrappingType.Anchor 浮动图像
* HorizontalPositionOffset: 设置图片相对于锚点的水平偏移量(以像素为单位)
* VerticalPositionOffset：设置图片相对于锚点的垂直偏移量（以像素为单位）
* BehindDoc: 控制图片是否显示在文档文字的后方
* AllowOverlap: 控制图片是否允许与其他图片或对象重叠


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







## 支持 : [Donate Link](https://miniexcel.github.io/)

<a href="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg"><img src="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg" alt="wechat" width="200px" height="300px">
</a>
<a href="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg">
    <img src="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg" alt="alipay" width="200px" height="300px"></a>



