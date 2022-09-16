<div align="center">
<p><a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/v/MiniWord.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/dt/MiniWord.svg" alt=""></a>  
<a href="https://github.com/mini-software/MiniWord" rel="nofollow"><img src="https://img.shields.io/github/stars/mini-software/MiniWord?logo=github" alt="GitHub stars"></a> 
<a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version"></a>
</p>
</div>


---

<div align="center">
 您的 <a href="https://github.com/mini-software/MiniWord">Star</a> 和 <a href="https://miniexcel.github.io">贊助</a> 可以讓 MiniWord 走更遠
</div>

---


## 介紹

MiniWord .NET Word模板引擎，藉由Word模板和數據簡單、快速生成文件。

![image](https://user-images.githubusercontent.com/12729184/190674408-12c03f86-31ea-4132-bb31-e2a793f8c40f.png)



## 標籤

### 文本

##### 代碼例子

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

##### 導出

![image](https://user-images.githubusercontent.com/12729184/190646113-04182d43-6b04-441d-911b-68de6af18039.png)

### 圖片

標籤值為 `MiniWordPicture` 類別

##### 代碼例子

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

標籤值為 `string[]` 或是 `IList<string>`類別

##### 代碼例子

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