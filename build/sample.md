
# 测试

## 段落

第一段。

第二段。

$$a + b = \frac{c}{d}$${#eq:eee}

$$\begin{matrix}1 & 2 \\ 3 & 4 \\\end{matrix}$$

公式引用：[@eq:eee]

## 图片

![图片描述](sample.jpg){#fig:fff width=100}

![图片置顶](sample.jpg){.frame width=100}

![图片置底](sample.jpg){.frame .frame-bottom width=100}

图片引用：[@fig:fff]

## 表格

- - -
1 2 3
4 5 6
- - -

## 书签和域

书签：[被引用的文字]{#test_ref}、[test_ref]{#test_ref_ref}

Ref域：`{Ref test_ref|Ref域}`{=field}、`{Ref {Ref test_ref_ref|wtf}|嵌套Ref域}`{=field}

Seq域：`{Seq count}`{=field}、`{Seq count}`{=field}、`{Seq count}`{=field}。

更新域，可以Ctrl+A全选后、F9更新。


# 换行

Soft break
test in
English.

中文
换行
测试。

The
mix
换行
测试
example.

混合
换行
test
example
。

# 引用一些文章

[@Centixkadon2020] 中做了一些微小的工作，减少了论文排版的工作量。

# 参考文献

:::{#refs}
:::
