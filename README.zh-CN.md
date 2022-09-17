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


## 介绍

MiniWord .NET Word模板引擎，藉由Word模板和数据简单、快速生成文件。

![image](https://user-images.githubusercontent.com/12729184/190674408-12c03f86-31ea-4132-bb31-e2a793f8c40f.png)



## 标签

### 文本

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

##### 效果

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

##### 效果

![image](https://user-images.githubusercontent.com/12729184/190645704-1f6405e9-71e3-45b9-aa99-2ba52e5e1519.png)

## 支持我 : [Donate Link](https://miniexcel.github.io/)

<a href="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg"><img src="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg" alt="wechat" width="200px" height="300px">
</a> 
<a href="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg">
    <img src="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg" alt="alipay" width="200px" height="300px"></a>

## 常见问题

### 模版字串没有生效

建议 `{{tag}}` 复制重新整串复制贴上，有时打字 word 在底层`{{}}`会被切开变成`<w:t>{</w:t><w:t>{<w:/t><w:t>Tag</w:t><w:t>}</w:t><w:t>}<w:/t>`如图片

![image](https://user-images.githubusercontent.com/12729184/190683025-fbf1bfa3-a34a-4af9-a8d3-30c6807d229c.png)



