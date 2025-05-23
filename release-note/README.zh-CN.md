## Release  Notes

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
 Your <a href="https://github.com/mini-software/MiniWord">Star</a> and <a href="https://miniexcel.github.io">Donate</a> can make MiniWord better 
</div>

---



### 0.9.1-0.9.2
- [Bug] Fix async (via @isdaniel)
- [Bug] Fix openxml version

### 0.9.0
- [New] Support async (@isdaniel)
- [New] Support new OpenXml to solve the problem of line wrapping to multiple lines #68 (via @ping9719)
- [New] Add support to conditional check for undefined tags #75 (via @bprucha)
- [New] Fix when a Split Tag Text element Inner Text start with " {" instead of "{" (via @hieplenet)
- [New] Extension: floating image, you can configure the image to float on the text (via @dessli)
- [New] Extension: Table supports Obj.objA.List.Prop1 rendering (via wangx036)
- [New] Extension: Common types support multi-level attribute rendering, such as {{Obj.A.B.C}} (via wangx036)
- [New] @foreach (via wangx036)
- [Bug] Fixed the problem that fonts with color, underline and other styles are normal in office, but not visible after opening with WPS (via @haozekang)
- [Bug] Remove all elements between @if~@endif, not paragraph (via wangx036)
- [Bug] Fix build, Fix tests, Update openxml (via @masterworgen)

### 0.8.0

- [New] 支持 new OpenXml to solve the problem of line wrapping to multiple lines #68 (via @ping9719)
- [New] 支持 to if statement inside foreach statement inside templates. Please refer to samples. (via @eynarhaji)
- [New] 变更 tags for if statements for single paragraph if statement {{if and endif}} inside templates. Please refer to samples. (via @eynarhaji)
- [Bug]  The table should be inserted at the template tag position instead of the last row #47 (via @itldg)

### 0.7.0
- [New] 支持 List inside List via `IEnumerable<MiniWordForeach>` and `{{foreach`/`endforeach}}` tags (via @eynarhaji)
- [New] 支持 @if statements inside templates (via @eynarhaji)
- [New] 支持 multiple color word by word (via @andy505050)

### 0.6.1
- [Bug] 修正系统不支持 `IEnumerable<MiniWordHyperLink>` (#39 via @isdaniel)

### 0.6.0
- [New] 支持 hyperLink  (#33 via @isdaniel)
- [New] 支持 custom font color and highlight color (#35 via impPDX)
- [New] 支持 2 level object parameter (#32 via @ping9719 , @shps951023)
- [Bug] 修正 Multiple tags format error (#37 via @shps951023)

### 0.5.0

- [New] 支持 object & dynamic parameter (#19 via @isdaniel )

### 0.4.0
- [New] 支持HeaderParts, FooterParts template
- [Bug] 修正multiple table generate problem #18

### 0.3.0
- [New] 支持 table 标签  #13
- [New] datetime format -> yyyy-MM-dd HH:mm:ss
- [Bug] fixed spliting template string like `<w:t>{</w:t><w:t>{<w:/t><w:t>Tag</w:t><w:t>}</w:t><w:t>}<w:/t>` problem #17

### 0.2.1

- [Bug] fixed mutiple tag System.InvalidOperationException: 'The parent of this element is null.' #13

### 0.2.0

- [Feature] 支持 array list string 生成多行 #11
- [Feature] 支持图片 #10 #3
- [Feature] 图片支持自定义 width 和 height #8
- [Feature] 支持多 breakline
- [Optimize] 删除 xmlns declaration #7

### 0.1.1
- [Bug] 修正 Fuzzy Regex replace similar key


### 0.1.0
- 基本 template 导出

### 0.0.0
- Init