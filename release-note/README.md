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

### 0.7.1
- [New] Add support to if statement inside foreach statement inside templates. Please refer to samples. (via @eynarhaji)
- [New] Change tags for if statements for single paragraph if statement {{if and endif}} inside templates. Please refer to samples. (via @eynarhaji)

### 0.7.0
- [New] Add support to List inside List via `IEnumerable<MiniWordForeach>` and `{{foreach`/`endforeach}}` tags (via @eynarhaji)
- [New] Add support to @if statements inside templates (via @eynarhaji)
- [New] Support multiple color word by word (via @andy505050)

### 0.6.1
- [Bug] Fixed system does not support `IEnumerable<MiniWordHyperLink>` (#39 via @isdaniel)

### 0.6.0
- [New] Support hyperLink  (#33 via @isdaniel)
- [New] Support custom font color and highlight color  (#35 via impPDX)
- [New] Support 2 level object parameter (#32 via @ping9719 , @shps951023)
- [Bug] Fix Multiple tags format error (#37 via @shps951023)

### 0.5.0
- [New] support object & dynamic parameter (#19 via @isdaniel )

  
### 0.4.0
- [New] support HeaderParts, FooterParts template
- [Bug] fixed multiple table generate problem #18

### 0.3.0
- [New] Support table generate  #13
- [New] datetime format -> yyyy-MM-dd HH:mm:ss
- [Bug] fixed spliting template string like `<w:t>{</w:t><w:t>{<w:/t><w:t>Tag</w:t><w:t>}</w:t><w:t>}<w:/t>` problem #17


### 0.2.1

- [Bug] fixed mutiple tag System.InvalidOperationException: 'The parent of this element is null.' #13



### 0.2.0

- [Feature] support array list string to generate multiple row #11
- [Feature] support image #10 #3
- [Feature] image support to custom width and height #8
- [Feature] support multiple breakline 
- [Optimize] Remove xmlns declaration #7

### 0.1.1

- [Bug] Fix Fuzzy Regex replace similar key

### 0.1.0
- [Feature] basic template export

### 0.0.0
- Init