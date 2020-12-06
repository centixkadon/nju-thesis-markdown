
-- configuration ---------------------------------------------------------------

local configuration = {
  "word_count", -- 字数统计
  "cover", -- 添加封面
  "numbering", -- 图表式编号及引用
  "section", -- 添加目录和分节符
  -- "frame", -- 简单文本框
  "bookmark", -- 书签
  "field", -- 域
  "softbreak", -- 段内换行
}

-- 公式和编号的列宽比例。如果公式编号格式不对，可以按需调整
local equation_widthpcts = pandoc.List{ 95, 5 }

-- -----------------------------------------------------------------------------



utf8.lower = pandoc.text.lower
utf8.upper = pandoc.text.upper
utf8.reverse = pandoc.text.reverse
-- utf8.len = pandoc.text.len
utf8.sub = pandoc.text.sub


function table.update(dict1, dict2)
  if dict2 then
    for key, value in pairs(dict2) do
      dict1[key] = value
    end
  end
end

function table.delete(dict)
  for key, value in pairs(dict) do
    if type(key) == "string" and (value == false or value == 0 or (type(value) == "string" and #value == 0)) then
      dict[key] = nil
    end
  end
end

function table.tostring(dict)
  local s = ""
  local t = false
  for key, value in pairs(dict) do
    if type(value) == "table" then
      value = table.tostring(value)
    elseif pandoc.List{ "nil", "boolean", "number", "string" }:includes(type(value)) then
      value = string.format("%q", value)
    else
      value = type(value)
    end

    if type(key) == "string" then
      t = true
      s = s..string.format(",%s=%s", key, value)
    else
      s = s..string.format(",%s", value)
    end
  end

  if t then
    return string.format("{%s}", string.sub(s,2))
  else
    return string.format("[%s]", string.sub(s,2))
  end
end

function table.rep(elem, n)
  local list = {}
  for i = 1,n do
    table.insert(list, elem)
  end
  return list
end


function pandoc.List:reduce(func, sum)
  sum = sum or 0
  for i, value in ipairs(self) do
    sum = func(sum, value)
  end
  return sum
end

function pandoc.List:erase(i, j)
  i = i or 1
  j = j or #self
  for k = math.min(j, #self), i, -1 do
    self:remove(k)
  end
end

function pandoc.CaptionLong(blocks)
  return { long=blocks }
end

function pandoc.CaptionShort(inlines)
  return { short=inlines }
end

function pandoc.ColSpec(width, alignment)
  return { alignment or pandoc.AlignDefault, width }
end

function pandoc.TableAttr(identifier, classes, attributes)
  return { identifier or "", classes or {}, attributes or {} }
end

function pandoc.TableCell(contents, row_span, col_span, alignment, attr)
  return {
    attr = attr or pandoc.TableAttr(),
    alignment = alignment or pandoc.AlignDefault,
    contents = contents,
    col_span = col_span or 1,
    row_span = row_span or 1,
  }
end

function pandoc.TableRow(cells, attr)
  return { attr or pandoc.TableAttr(), cells }
end

function pandoc.TableHead(rows, attr)
  return { attr or pandoc.TableAttr(), rows }
end

function pandoc.TableBody(body, row_head_columns, head, attr)
  return {
    attr = attr or pandoc.TableAttr(),
    body = body,
    head = head or {},
    row_head_columns = row_head_columns or 0,
  }
end

function pandoc.TableFoot(rows, attr)
  return { attr or pandoc.TableAttr(), rows or {} }
end


function pandoctostring(elem)
  if elem.tag == "Space" then
    return " "
  elseif elem.tag == "Str" then
    return elem.text
  elseif elem.tag == "Quoted" then
    local s = ""
    for i = 1, #elem.content do
      s = s..pandoctostring(elem.content[i])
    end
    if elem.quotetype == "SingleQuote" then return "'"..s.."'" end
    if elem.quotetype == "DoubleQuote" then return '"'..s..'"' end
    return s
  end
  return ""
end

function inlinestoattr(inlines)
  local s = ""
  for i = #inlines, 1, -1 do
    s = pandoctostring(inlines[i])..s

    if string.sub(s,1,1) == "{" then
      if string.sub(s,-1) == "}" then

        s = s:sub(2,-2)
        local identifier = s:match("#([%w-_]+)") or ""
        local classes = {}
        local attributes = {}
        for class in s:gmatch("%.([%w-_]+)") do table.insert(classes, class) end
        for key, value in s:gmatch("([%w-_]+)=([%w-_]+)") do attributes[key] = value end
        for key, value in s:gmatch('([%w-_]+)="([%w%s-_]+)"') do attributes[key] = value end

        inlines:erase(i)
        for i = i-1, 1, -1 do
          if inlines[i].tag == "Space" then
            inlines:remove()
          else
            break
          end
        end

        return pandoc.Attr(identifier, classes, attributes)
      end

      break
    end
  end

end

function printobject(obj, indent, prefix)
  if indent == nil then
    indent = 0
  end
  local s = string.rep(" ", indent)
  if prefix ~= nil then
    s = s..prefix.." = "
  end
  if type(obj) == "table" then
    s = s.."{"
    print(s)
    if obj.tag then
      printobject(obj.tag, indent + 2, "tag")
    end
    for key, value in pairs(obj) do
      if type(key) == "string" then
        if key ~= "tag" then
          printobject(value, indent + 2, key)
        end
      else
        printobject(value, indent + 2)
      end
    end
    print(string.rep(" ", indent).."},")
  elseif pandoc.List{ "nil", "boolean", "number", "string" }:includes(type(obj)) then
    print(string.format("%s%q,", s, obj))
  else
    print(string.format("%s%q,", s, type(obj)))
  end
end



if pandoc.List{ "docx", "json" }:includes(FORMAT) then

io.output("thesis.log")
configuration = pandoc.List(configuration)
local filters = pandoc.List{}

local raws = {
  tab = pandoc.RawInline("openxml", '<w:r><w:tab /></w:r>'),
  page = pandoc.RawInline("openxml", '<w:br w:type="page" />'),
}

if configuration:includes("word_count") then
  local word_number = 0

  filters:insert{
    Str = function (elem)
      local t = true
      for c in string.gmatch(elem.text, utf8.charpattern) do
        local tt = #c > 1
        if t or tt then
          word_number = word_number + 1
        end
        t = tt
      end
    end,

    Pandoc = function (elem)
      io.write(string.format("word number: %q\n", word_number))
    end,
  }
end

if configuration:includes("cover") then
  local metavalues = {}
  filters:insert{
    Meta = function (elem)
      pandoc.List{ "cover-image", "cover-degree" }:map(function (key)
        if elem[key] then
          local s = ""
          pandoc.List(elem[key]):map(function (x) s = s..pandoctostring(x) end)
          metavalues[key] = s
          elem[key] = nil
        end
      end)

      pandoc.List{ "title", "author", "specialization", "major", "mentor", "date", "number", "grade", "keywords", "title-en", "author-en", "specialization-en", "mentor-en", "keywords-en" }:map(function (key)
        if elem[key] then
          if elem[key].tag == "MetaInlines" then metavalues[key] = pandoc.List(elem[key]) end
          elem[key] = nil
        end
      end)

      pandoc.List{ "title-cover", "author-cover", "specialization-cover", "major-cover", "mentor-cover", "abstract", "abstract-en" }:map(function (key)
        if elem[key] then
          if elem[key].tag == "MetaBlocks" then metavalues[key] = pandoc.List(elem[key]) end
          if elem[key].tag == "MetaInlines" then metavalues[key] = pandoc.List{ pandoc.Para(pandoc.List(elem[key])) } end
          elem[key] = nil
        end
      end)

      pandoc.List{ { "msg-sign", string.rep(" ", 40).."（签字）" } }:map(function (x)
        metavalues[x[1]] = pandoc.List{ pandoc.Str(x[2]) }
      end)

      return elem

    end,

  }

  local cover_filters = {
    bachelor = {},

    master = {
      Pandoc = function (elem)
        local labelcell = function (str, style, width)
          return pandoc.TableCell({
            pandoc.RawBlock("openxml", string.format('<w:tcPr><w:tcMar><w:right w:w="0" w:type="dxa" /></w:tcMar><w:tcW w:w="%q" w:type="pct" /></w:tcPr><w:p><w:pPr><w:pStyle w:val="%s" /><w:jc w:val="distribute" /></w:pPr><w:r><w:rPr><w:spacing w:val="-200" /></w:rPr><w:t>%s</w:t></w:r><w:r><w:t>%s</w:t></w:r></w:p>', width, string.gsub(style, "%s", ""), utf8.sub(str,1,-2), utf8.sub(str,-1)))
          })
        end
        local textcell = function (list, width)
          local textpr = pandoc.RawBlock("openxml", string.format('<w:tcPr><w:tcMar><w:left w:w="0" w:type="dxa" /></w:tcMar><w:tcW w:w="%q" w:type="pct" /></w:tcPr>', width))
          return pandoc.TableCell({ textpr, pandoc.Para({ raws.tab, pandoc.Span(list), raws.tab }) })
        end

        local message = function (tablepairs, tablesep, widthpcts, styles)
          local rows = pandoc.List{}
          for _, list in ipairs(tablepairs) do
            local label, text = table.unpack(list)
            for _, value in ipairs(metavalues[text.."-cover"] or { pandoc.Para(metavalues[text]) }) do
              rows:insert(pandoc.TableRow({ labelcell(label, styles[1], widthpcts[1]), pandoc.TableCell({
                pandoc.RawBlock("openxml", string.format('<w:tcPr><w:tcMar><w:left w:w="0" w:type="dxa" /><w:right w:w="0" w:type="dxa" /></w:tcMar><w:tcW w:w="%q" w:type="pct" /></w:tcPr><w:p><w:pPr><w:pStyle w:val="%s" /></w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>', widthpcts[2], string.gsub(styles[1], "%s", ""),tablesep))
              }), textcell(value.content, widthpcts[3]) }))
              label = ""
            end
          end
          return pandoc.Div(pandoc.Table(
            pandoc.CaptionLong({}),
            pandoc.List(widthpcts):map(function (x) return pandoc.ColSpec(x / 100) end),
            pandoc.TableHead({}),
            { pandoc.TableBody(rows) },
            pandoc.TableFoot({})
          ), { ["custom-style"]=styles[2] })
        end

        local cover = pandoc.List{
          pandoc.Para(pandoc.Str("")),
          pandoc.Div(pandoc.Para(pandoc.Image({}, metavalues["cover-image"], "", { width=187 })), { ["custom-style"]="Captioned Figure" }),
          pandoc.Div({
            pandoc.Para({ pandoc.Str("研  究  生  毕  业  论  文") }),
            pandoc.Para({ pandoc.Str("（申请硕士学位）") }),
          }, { ["custom-style"]="Title" }),
          pandoc.Para(table.rep(pandoc.LineBreak(), 3)),
          message({ { "论文题目", "title" }, { "作者姓名", "author" }, { "学科、专业名称", "specialization" }, { "研究方向", "major" }, { "指导老师", "mentor" } }, "", { 23, 5, 60 }, { "Front Cover Label", "Front Cover Text" }),
          pandoc.Div(pandoc.Para(table.rep(pandoc.LineBreak(), 5 - #metavalues["title-cover"])), { ["custom-style"]="Front Cover Text" }),
          pandoc.Div({ pandoc.Para({ pandoc.Span(metavalues["date"]), raws.page }) }, { ["custom-style"]="Date" }),

          pandoc.Para(table.rep(pandoc.LineBreak(), 26)),
          message({ { "学号", "number" }, { "论文答辩日期", "date" }, { "指导教师", "msg-sign" } }, "：", { 23, 5, 72 }, { "Inside Front Cover Label", "Inside Front Cover Text" }),
          pandoc.Para({ raws.page }),

          pandoc.Div({
            pandoc.Para({ pandoc.Str("南京大学研究生毕业论文中文摘要首页用纸") }),
          }, { ["custom-style"]="Abstract Title" }),
          pandoc.Div({
            pandoc.Para({ pandoc.Span(pandoc.Str("毕业论文题目："), { ["custom-style"]="Abstract Char" }), pandoc.Str("　"), pandoc.Span(metavalues["title"]), pandoc.Str("\t") }),
            pandoc.Para({
              pandoc.Str("　"), pandoc.Span(metavalues["specialization"]), pandoc.Str("　"), pandoc.Span(pandoc.Str("专业"), { ["custom-style"]="Abstract Char" }),
              pandoc.Str("　"), pandoc.Span(metavalues["grade"]), pandoc.Str("　"), pandoc.Span(pandoc.Str("级硕士生姓名："), { ["custom-style"]="Abstract Char" }),
              pandoc.Str("　"), pandoc.Span(metavalues["author"]), pandoc.Str("\t"),
            }),
            pandoc.Para({ pandoc.Span(pandoc.Str("指导教师（姓名、职称）："), { ["custom-style"]="Abstract Char" }), pandoc.Str("　"), pandoc.Span(metavalues["mentor"]), pandoc.Str("\t") }),
            pandoc.Para(pandoc.Span(pandoc.Str("摘要："), { ["custom-style"]="Abstract Char" })),
          }, { ["custom-style"]="Abstract Underline Message" }),
          pandoc.Div(metavalues["abstract"], { ["custom-style"]="Abstract" }),
          pandoc.Div(pandoc.Para({ pandoc.Str("关键词："), pandoc.Span(metavalues["keywords"]) }), { ["custom-style"]="Keywords" }),
          pandoc.Para({ raws.page }),

          pandoc.Div({
            pandoc.Para({ pandoc.Str("南京大学研究生毕业论文英文摘要首页用纸") }),
          }, { ["custom-style"]="Abstract Title" }),
          pandoc.Div({
            pandoc.Para({ pandoc.Str("THESIS: "), pandoc.Span(metavalues["title-en"]) }),
            pandoc.Para({ pandoc.Str("SPECIALIZATION: "), pandoc.Span(metavalues["specialization-en"]) }),
            pandoc.Para({ pandoc.Str("POSTGRADUATE: "), pandoc.Span(metavalues["author-en"]) }),
            pandoc.Para({ pandoc.Str("MENTOR: "), pandoc.Span(metavalues["mentor-en"]) }),
            pandoc.Para(pandoc.Str("ABSTRACT: ")),
          }, { ["custom-style"]="Abstract Message" }),
          pandoc.Div(metavalues["abstract-en"], { ["custom-style"]="Abstract" }),
          pandoc.Div(pandoc.Para({ pandoc.Str("KEYWORDS: "), pandoc.Span(metavalues["keywords-en"]) }), { ["custom-style"]="Keywords" }),
        }

        cover:extend(elem.blocks)
        elem.blocks = cover
        return elem
      end,

    },

    phd = {},
  }

  local nilfunc = function (elem) end
  filters:insert{
    Pandoc = function (elem) return ((cover_filters[metavalues["cover-degree"]] or {}).Pandoc or nilfunc)(elem) end,
  }
end

if configuration:includes("numbering") then
  local number = { 0, 0, 0, 0, 0, 0, fig=0, tbl=0, eq=0 }
  local numbers = {}

  filters:insert{
    Header = function (elem)
      number[elem.level] = number[elem.level] + 1
      for i = elem.level+1, #number do
        number[i] = 0
      end
      if elem.level <= 1 then
        number.fig = 0
        number.lst = 0
        number.eq = 0
      end
    end,

    Para = function (elem)
      if elem.content[1].tag == "Math" and elem.content[1].mathtype == "DisplayMath" then
        number.eq = number.eq + 1

        local widthpcts = equation_widthpcts
        local widths = widthpcts:map(function (x) return x / 100 end)
        local tcpr = widthpcts:map(function (x) return pandoc.RawBlock("openxml", string.format('<w:tcPr><w:tcW w:w="%q" w:type="pct"/></w:tcPr>', x)) end)

        local attr = inlinestoattr(elem.content)
        local eqno = {
          pandoc.Str("("),
          pandoc.RawInline("field", string.format('{={Section \\* MergeFormat|}-1|%q}', number[1])),
          pandoc.Str("."),
          pandoc.RawInline("field", string.format("{Seq equations|%q}", number.eq)),
          pandoc.Str(")"),
        }
        if attr and attr.identifier and attr.identifier:sub(1,3) == "eq-" then
          if numbers[attr.identifier] then
            io.write(string.format("bookmark exist: %q\n", attr.identifier))
          else
            numbers[attr.identifier] = string.format("(%s.%s)", number[1], number.eq)
          end
          eqno = pandoc.Span(eqno, { id=attr.identifier })
        end
        return { pandoc.Table(
          pandoc.CaptionLong({}),
          { pandoc.ColSpec(widths[1], pandoc.AlignLeft), pandoc.ColSpec(widths[2], pandoc.AlignRight) },
          pandoc.TableHead({}),
          { pandoc.TableBody({ pandoc.TableRow({
            pandoc.TableCell({ tcpr[1], pandoc.Div(elem, { ["custom-style"]="Equation" }) }),
            pandoc.TableCell({ tcpr[2], pandoc.Div(pandoc.Para(eqno), { ["custom-style"]="Equation Caption" }) }),
          }) }) },
          pandoc.TableFoot({})
        ) }

      elseif elem.content[1].tag == "Image" and not elem.content[2] then
        local image = elem.content[1]
        if #image.caption > 0 then
          number.fig = number.fig + 1

          if image.identifier and image.identifier:sub(1,4) == "fig-" then
            numbers[image.identifier] = string.format("图%s.%s", number[1], number.fig)
            image.caption = { pandoc.Span(image.caption, { id=image.identifier }) }
            image.identifier = ""
          end
        end

        return elem
      end
    end,

    Table = function (elem)
      if elem.caption.long and #elem.caption.long > 0 then
        number.tbl = number.tbl + 1

        local plain = elem.caption.long[1]
        local attr = inlinestoattr(plain.content)
        if attr and attr.identifier and attr.identifier:sub(1,4) == "tbl-" then
          numbers[attr.identifier] = string.format("表%s.%s", number[1], number.tbl)
          plain.content = { pandoc.Span(plain.content, { id=attr.identifier }) }
        end
      end

      return elem
    end,

  }

  filters:insert{
    Cite = function (elem)
      local identifier = elem.citations[1].id
      if string.sub(identifier, 1, 3) == "eq-" then
        return {
          pandoc.Str("式"),
          pandoc.RawInline("field", string.format("{Ref %s|%s}", identifier, numbers[identifier] or "Error! Equation not defined")),
        }
      elseif string.sub(identifier, 1, 4) == "fig-" then
        return { pandoc.RawInline("field", string.format("{Ref %s \\r|%s}", identifier, numbers[identifier] or "Error! Figure not defined")) }
      elseif string.sub(identifier, 1, 4) == "tbl-" then
        return { pandoc.RawInline("field", string.format("{Ref %s \\r|%s}", identifier, numbers[identifier] or "Error! Table not defined")) }
      end
    end,
  }

end

if configuration:includes("section") then
  local number = { 0, 0, 0, 0, 0, 0 }
  local sectionProperties = {
    [[
      <w:p><w:r><w:br w:type="page" /></w:r></w:p><w:sdt>
        <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents" /><w:docPartUnique /></w:docPartObj></w:sdtPr>
        <w:sdtContent>
          <w:p><w:pPr><w:pStyle w:val="TOCHeading" /></w:pPr><w:r><w:t xml:space="preserve">目录</w:t></w:r></w:p>
          <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true" />
            <w:instrText xml:space="preserve">TOC \o "1-3" \h \z \u</w:instrText>
          <w:fldChar w:fldCharType="separate" /><w:fldChar w:fldCharType="end" /></w:r></w:p>
        </w:sdtContent>
      </w:sdt><w:p><w:pPr><w:sectPr>
        <w:footerReference w:type="default" r:id="rId10" />
        <w:pgSz w:w="11907" w:h="16840" />
        <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0" />
        <w:pgNumType w:fmt="upperRoman" w:start="0" />
        <w:cols w:space="720" />
        <w:titlePg />
      </w:sectPr></w:pPr></w:p>
    ]],
    [[
      <w:p><w:pPr><w:sectPr>
        <w:headerReference w:type="default" r:id="rId9" />
        <w:footerReference w:type="default" r:id="rId10" />
        <w:pgSz w:w="11907" w:h="16840" />
        <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0" />
        <w:pgNumType w:fmt="demical" w:start="1" />
        <w:cols w:space="720" />
      </w:sectPr></w:pPr></w:p>
    ]],
    [[
      <w:p><w:pPr><w:sectPr>
        <w:headerReference w:type="default" r:id="rId9" />
        <w:footerReference w:type="default" r:id="rId10" />
        <w:pgSz w:w="11907" w:h="16840" />
        <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0" />
        <w:cols w:space="720" />
      </w:sectPr></w:pPr></w:p>
    ]],
  }

  filters:insert{
    Header = function (elem)
      number[elem.level] = number[elem.level] + 1
      for i = elem.level+1, #number do
        number[i] = 0
      end
      if elem.level == 1 then
        chapter = number[elem.level]
        elem.content = { pandoc.Span(elem.content, { id=string.format("chapter-%s", chapter) }) }
        return { pandoc.RawBlock("openxml", sectionProperties[math.min(chapter, #sectionProperties)]), elem }
      end
    end,

  }

  filters:insert{
    Pandoc = function (elem)
      if number[1] == 1 then
        elem.blocks:insert(pandoc.RawBlock("openxml", '<w:sectPr><w:headerReference w:type="default" r:id="rId9" /><w:pgNumType w:fmt="demical" w:start="1" /></w:sectPr>'))
        return elem
      end
    end,

  }

end

if configuration:includes("frame") then
  filters:insert{
    Div = function (elem)
      if elem.classes:includes("frame") then
        local properties = {
          wrap="around", -- around, none
          w="", -- [auto], 0, ...
          hRule="auto", -- [auto], atLeast, exact
          h="", -- 0, ...
          xAlign="left", -- [x], left, right, center, inside, outside
          x="", -- [0], ...
          hAnchor="margin", -- [margin], page, text
          yAlign="top", -- [y], top, bottom, center, inside, outside
          y="", -- [0], ...
          vAnchor="margin", -- margin, page, text (ignore yAlign)
          anchorLock="1", -- [0], 1
        }
        if elem.classes:includes("frame-left") then table.update(properties, { xAlign="left" }) end
        if elem.classes:includes("frame-right") then table.update(properties, { xAlign="right" }) end
        if elem.classes:includes("frame-top") then table.update(properties, { yAlign="top" }) end
        if elem.classes:includes("frame-bottom") then table.update(properties, { yAlign="bottom" }) end
        table.update(properties, elem.attributes)
        table.delete(properties)

        local frame = '<w:pPr><w:framePr '
        for key, value in pairs(properties) do
          frame = frame..string.format("w:%s=%q ", key, value)
        end
        frame = frame..'/></w:pPr>'
        prefix = pandoc.RawInline("openxml", frame)

        for i, block in pairs(elem.content) do
          local t = true
          local filter = function (elem)
            if t or (elem.tag == "Image" and pandoc.utils.equals(elem.caption[1], prefix)) then
              t = false
              return { prefix, elem }
            end
          end
          elem.content[i] = pandoc.walk_block(block, { Inline=filter })
        end
        return elem
      end

    end,

  }

end

if configuration:includes("bookmark") then
  filters:insert{
    RawInline = function (elem)
      if elem.format:find("bookmark") then
        local text = ''
        if ("bookmarkStart"):find(elem.format) then text = text..string.format('<w:bookmarkStart w:id=%q w:name=%q/>', elem.text, elem.text) end
        if ("bookmarkEnd"):find(elem.format) then text = text..string.format('<w:bookmarkEnd w:id=%q/>', elem.text) end
        return pandoc.RawInline("openxml", text)
      end
    end,

  }

end

if configuration:includes("field") then
  filters:insert{
    RawInline = function (elem)
      if elem.format:find("field") then
        local text = '<w:r>'
        local instr = ""
        for c in string.gmatch(elem.text, utf8.charpattern) do
          if c == "{" then
            if #instr > 0 then
                text = text..string.format('<w:instrText xml:space="preserve">%s</w:instrText>', instr)
            end
            text, instr = text..'<w:fldChar w:fldCharType="begin"/>', ""
          elseif c == "}" then
            local textstr = "Field"
            if #instr > 0 then
              for j = utf8.len(instr),1,-1 do
                if utf8.sub(instr,j,j) == "|" then
                  instr, textstr = utf8.sub(instr, 1, j - 1), utf8.sub(instr, j + 1)
                  break
                end
              end
              if #instr > 0 then text = text..string.format('<w:instrText xml:space="preserve">%s</w:instrText>', instr) end
            end
            text = text..'<w:fldChar w:fldCharType="separate"/>'
            if #textstr > 0 then text = text..string.format('<w:t>%s</w:t>', textstr) end
            text, instr = text..'<w:fldChar w:fldCharType="end"/>', ""
          else
            instr = instr..c
          end
        end
        return pandoc.RawInline("openxml", text..'</w:r>')

      end
    end,

  }

end

if configuration:includes("softbreak") then
  filters:insert{
    Inlines = function (inlines)
      for i = #inlines-1, 2, -1 do
        if inlines[i].tag == "SoftBreak" then
          local l = inlines[i-1]
          if l.tag ~= "Str" then
            l = nil
            pandoc.walk_inline(inlines[i-1], { Str = function (elem) l = elem end })
          end

          if l and #utf8.sub(l.text, -1) > 1 then
            inlines:remove(i)
          else
            local r = inlines[i+1]
            if r.tag ~= "Str" then
              r = nil
              pandoc.walk_inline(inlines[i+1], { Str = function (elem) if not r then r = elem end end })
            end

            if r and #utf8.sub(r.text, 1, 1) > 1 then
              inlines:remove(i)
            end
          end
        end
      end
      return inlines
    end,

  }

end

if configuration:includes("bookmark") then
  filters:insert{
    Div = function (elem)
      if utf8.len(elem.identifier) > 40 then
        io.write(string.format("bookmark too long: %q\n", elem.identifier))
      end
    end,

    Span = function (elem)
      if utf8.len(elem.identifier) > 40 then
        io.write(string.format("bookmark too long: %q\n", elem.identifier))
      end
    end,

  }

end

return filters

end -- docx json
