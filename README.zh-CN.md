<div align="center">
<p><a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/v/MiniWord.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/nuget/dt/MiniWord.svg" alt=""></a>  
<a href="https://github.com/mini-software/MiniWord" rel="nofollow"><img src="https://img.shields.io/github/stars/mini-software/MiniWord?logo=github" alt="GitHub stars"></a> 
<a href="https://www.nuget.org/packages/MiniWord"><img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version"></a>
</p>
</div>

---

<div align="center">
 您的 <a href="https://github.com/mini-software/MiniWord">Star</a> 和 <a href="https://miniexcel.github.io">赞助</a> 可以让 MiniWord 走更远
</div>
---


### 介绍

MiniWord 简单 Word 模版导出+填充数据工具。

### 基本模版导出

```csharp
			var value = new Dictionary<string, object>()
			{
				["Company_Name"] = "MiniSofteware",
				["Name"] = "Jack",
				["CreateDate"] = new DateTime(2021, 01, 01),
				["VIP"] = true,
				["Points"] = 123,
				["APP"] = "Demo APP",
			};
			MiniSoftware.MiniWord.SaveAsByTemplate(path, templatePath, value);
```

模版:

![image](https://user-images.githubusercontent.com/12729184/189614577-ac22d47c-30d5-4db5-9299-09f07211f1bf.png)

结果:

![image](https://user-images.githubusercontent.com/12729184/189612248-dd9381de-bbb8-4c72-adec-ac8982f60f96.png)