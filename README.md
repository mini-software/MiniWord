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
 Your <a href="https://github.com/mini-software/MiniWord">Star</a> and <a href="https://miniexcel.github.io">dotnate</a> can make MiniWord better
</div>

---

## Introduction

MiniWord is an easy and effective .NET Word Template library.

![image](https://user-images.githubusercontent.com/12729184/190674408-12c03f86-31ea-4132-bb31-e2a793f8c40f.png)



## Tags

### Text

##### Example

```csharp
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
```

##### Result

![image](https://user-images.githubusercontent.com/12729184/190646113-04182d43-6b04-441d-911b-68de6af18039.png)

### Image

标签值为 `MiniWordPicture` 类别

##### Example

```csharp
var value = new Dictionary<string, object>()
{
    ["Logo"] = new MiniWordPicture() { Path= PathHelper.GetFile("DemoLogo.png"), Width= 180, Height= 180 }
};
MiniWord.SaveAsByTemplate(path, templatePath, value);
```



##### Template

![image](https://user-images.githubusercontent.com/12729184/190647953-6f9da393-e666-4658-a56d-b3a7f13c0ea1.png)

##### Result

![image](https://user-images.githubusercontent.com/12729184/190648179-30258d82-723d-4266-b711-43f132d1842d.png)

### List

tag value is `string[]` or `IList<string>` type

##### Example

```csharp
var value = new Dictionary<string, object>()
{
    ["managers"] = new[] { "Jack" ,"Alan"},
    ["employees"] = new[] { "Mike" ,"Henry"},
};
MiniWord.SaveAsByTemplate(path, templatePath, value);
```

Template

![image](https://user-images.githubusercontent.com/12729184/190645513-230c54f3-d38f-47af-b844-0c8c1eff2f52.png)

##### Result

![image](https://user-images.githubusercontent.com/12729184/190645704-1f6405e9-71e3-45b9-aa99-2ba52e5e1519.png)



## Support Me : [Donate Link](https://miniexcel.github.io/)

<a href="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg"><img src="https://user-images.githubusercontent.com/12729184/158003727-ca348041-5e59-44bc-a694-f400777e0252.jpg" alt="wechat" width="200px" height="300px">
</a>
<a href="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg">
    <img src="https://user-images.githubusercontent.com/12729184/158003731-6d132872-19c3-4840-b1af-97aa22f9bf4b.jpg" alt="alipay" width="200px" height="300px"></a>



## Questions 

### {{Tag}} doesn't work

Suggest to recopy and paste `{{tag}}` , sometime type tag in word`{{}}`will split to <w:t>{</w:t><w:t>{<w:/t><w:t>Tag</w:t><w:t>}</w:t><w:t>}<w:/t>` like image:

![image](https://user-images.githubusercontent.com/12729184/190683025-fbf1bfa3-a34a-4af9-a8d3-30c6807d229c.png)