
# nju-thesis-markdown

Life is short, you need Markdown.

本项目提供了一个基于 [Pandoc](https://www.pandoc.org/) 导出 docx 南京大学学位论文的 Lua Filter 。该模板虽暂不严格遵循某些格式要求，但想必可以满足学士和硕士学位论文的需求，可帮助没有 LaTeX 基础且没有强迫症的南京大学同学们极其方便地排版出够用的学位论文。

## 功能特色

- 可排版硕士、学士学位论文；
- 导出的 docx 文件用书签和域来引用，插入图、表、公式导致的引用文字变化可以直接更新；
- 导出的 docx 文件可以供不使用 LaTeX 的导师修改；
- 功能不足的地方可以导出 docx 文件后用 Microsoft Word 补足。

## 参考的格式要求

- [2020届本科毕业论文工作手册](https://jw.nju.edu.cn/22/24/c24683a467492/page.htm)
- [博士（硕士）学位论文编写格式规定（试行）](https://grawww.nju.edu.cn/ec/4e/c1040a60494/page.htm)

## 项目进展

- [ ] 前置部分
  - [ ] 封面
    - [ ] 图片
    - [ ] 〇〇毕业〇〇
    - [ ] 论文信息
  - [ ] 前言
  - [ ] 中英文摘要、关键词
    - [ ] 摘要标题
    - [ ] 摘要信息
    - [ ] 摘要
    - [ ] 关键词
  - [x] 目录、目次页
  - [x] 插图和附表清单（使用表格实现）
  - [x] 注释说明汇集表（使用表格实现）
  - [x] 规范要求
    - [x] 前置部分页码
- [x] 主体部分
  - [x] 序号
  - [x] 图、表、公式
    - [x] 编号
    - [x] 图、公式引用
    - [ ] 表引用
  - [x] 参考文献
    - [x] 规范格式
    - [x] 导出 docx 后用其它文献管理软件（如 Zotero 、 EndNote ）
  - [x] 科研成果
  - [x] 致谢
  - [x] 规范要求
    - [x] 字体、字号
    - [x] 页眉
    - [x] 正文页码
- [x] 附录部分（非必要）
  - [x] 附录页码（接正文页码）
- [ ] 结尾部分（非必要）
- [x] 杂项
  - [x] 域支持
  - [x] 中文之间的换行不添加空格（中英文之间的换行也不添加，英文之间的换行添加）
  - [x] 字数统计

## 简单开始

在 [Releases · jgm/pandoc](https://github.com/jgm/pandoc/releases) 下载最新的 pandoc 二进制文件。后缀为 -x86_64.zip 的是供 Windows 的免安装版本，其余同理。下载后解压。

在本项目用 Clone or download 下载最新的工程文件，解压到目录 nju-thesis-markdown

Windows 下打开 powershell 或 cmd 并进入目录 nju-thesis-markdown/build ，运行

```
/path/to/pandoc.exe --lua-filter ../src/nju-thesis.lua --bibliography sample.bib --filter pandoc-citeproc sample.md --reference-doc nju-thesis-reference.docx --output sample.docx
```

如需要导出 docx 文件再自行添加参考文献，则运行

```
/path/to/pandoc.exe --lua-filter ../src/nju-thesis.lua sample.md --reference-doc nju-thesis-reference.docx --output build/sample.docx
```

## 一些技巧

### 排版更改

导出 docx 文件的字体样式、布局排版、小节编号等等都可以在 build/nju-thesis-reference.docx 中更改。打开该文件，修改并更新对应的样式，然后保存即可。需要注意的是，由于项目不够完善，直接修改页眉页脚可能会导致最终导出 docx 文件出错。如有需求，可以在 src/nju-thesis-reference 中更改（此处需要学习 OOXML ），然后将该文件夹内的内容压缩成 zip 压缩文件 build/nju-thesis-reference.docx （注意不能包括 src/nju-thesis-reference 文件夹本身）。南大同学也可以直接提 [issue](https://github.com/centixkadon/nju-thesis-markdown/issues) 。

### 公式输入

Pandoc 直接支持 TeX 格式的公式，示例见 [Pandoc - Math Demos](https://pandoc.org/demo/math.text) 。如对 TeX 不熟悉，可在 Microsoft Word 中用自带的公式编辑器（快捷键 Alt+=）输入，保存成 math.docx 后运行以下命令查看 Markdown 表示方法：

```
/path/to/pandoc.exe math.docx --to markdown
```

### 其它

详见 src/sample.md 。包括图、表、公式、对图表式编号的引用、书签和域、引用文章和参考文献。

## 进阶使用

首先，心中默念三遍：“ Pandoc 很厉害，自己想实现的功能都能基于它实现。”

### Pandoc

Pandoc 的使用可以参考 [Pandoc - Pandoc User's Guide](https://www.pandoc.org/MANUAL.html) 。上一节使用的 Pandoc 命令行参数分别是以下作用：

| 命令行参数          | 作用                                           |
|:--------------------|:-----------------------------------------------|
| --lua-filter xxx    | 指定 Lua Filter 文件                           |
| --filter xxx        | 指定 Filter 程序                               |
|                     | 注意： --lua-filter 和 --filter 按顺序依次执行 |
| --bibliography xxx  | 指定参考文献文件                               |
| --reference-doc xxx | 指定格式文件                                   |

更多命令行参数及其用法，参见 [Options](https://www.pandoc.org/MANUAL.html#options) 一章。

### Pandoc's Markdown

编写论文所需的 Markdown 在原有的语法基础上添加了一些 Pandoc 特有的语法，可以参考 [Pandoc - Pandoc User's Guide](https://www.pandoc.org/MANUAL.html) 中 [Pandoc's Markdown](https://www.pandoc.org/MANUAL.html#pandocs-markdown) 一章。

相关示例见 [Pandoc - Demos](https://pandoc.org/demos.html)

### Lua Filter

如果在 Pandoc's Markdown 的基础上想减少一些重复劳动，可以参考 [Pandoc - Pandoc Lua Filters](https://pandoc.org/lua-filters.html) 修改或重新创建 Lua Filter 。比如，章节自动编号、字数统计等功能可以通过 Lua Filter 实现。

有关 [Lua](http://www.lua.org/home.html) 语法，可以参考 [Programming in Lua](http://www.lua.org/pil/) 或 [Lua: reference manuals](https://www.lua.org/manual/) 。

如果不想学习 Lua ，可以参考 [Pandoc - Pandoc filters](https://www.pandoc.org/filters.html) 用其它语言创建 Filter 。

### OOXML

OOXML (Office Open XML) 是 Microsoft 开发的、基于 XML (Extensible Markup Language) 的 zip 压缩文件格式。将 docx 文件后缀名改成 zip 并解压，就能得到里面的 XML 文件。 docx 文件的强大相信经常使用 Microsoft Word 的同学们可以从 bug 中体会到。这些 bug 往往仅仅是为了更方便的功能存储了冗余信息、图形用户界面和 XML 信息交互出错、甚至是压缩出了问题等等鸡毛蒜皮的问题导致的（~~我猜的~~）。那么怎样才能既使用这些强大的功能，又不出奇奇怪怪的问题呢？当然是直接编写 XML 文件再转换成 docx 啦。但是 OOXML 实在是难，不是人学的。这也不是问题， XML 是一门标记语言，自然可以由其它标记语言转换得到啦。 Markdown 就是这门简单得几乎不能再简单的标记语言。

Markdown 作为一门从命名就可以看出与 Markup Language 针锋相对的标记语言，因为简单方便，所以表达的信息是有极限的。好在 Pandoc's Markdown 提供了 raw_attribute 的扩展，可以在 Markdown 中写目标格式的标记语言。比如章节前自动换页、域等功能就是通过 Lua Filter 和 OOXML 实现的。

有关 OOXML 语法，官方的网站没有找到。推荐 [Office Open XML - What is OOXML?](http://officeopenxml.com/) ，虽然看起来不官方，但是非常全面详细，目前想找的都能找到。
