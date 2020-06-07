
-- configuration ---------------------------------------------------------------

local configuration = {
  "word_count",
  "numbering",
  "section",
  "utility",
  "bookmark_length",
}

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


function pandoc.List:reduce(func, sum)
  if not sum then
    sum = 0
  end
  for i, value in ipairs(self) do
    sum = func(sum, value)
  end
  return sum
end



if pandoc.List{ "docx", "json" }:includes(FORMAT) then

configuration = pandoc.List(configuration)
io.output("thesis.log")
local filters = pandoc.List{}

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

if configuration:includes("numbering") then
  local number = { 0, 0, 0, 0, 0, 0, fig=0, lst=0, eq=0 }
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

        local aligns = { "AlignLeft", "AlignRight" }
        local widthpcts = pandoc.List{ 95, 5 }
        local width = widthpcts:reduce(function (sum, x) return sum + x end)
        local widths = widthpcts:map(function (value) return value / width end)

        local formatstring = '<w:tcPr><w:tcW w:w="%q" w:type="pct"/></w:tcPr>'
        local tcpr = widthpcts:map(function (x) return pandoc.RawBlock("openxml", string.format(formatstring, x)) end)
        local math = pandoc.Para({ elem.content[1] })
        local eqno = {
          pandoc.Str("("),
          pandoc.RawInline("field", string.format('{={Section \\* MergeFormat|}-1|%q}', number[1])),
          pandoc.Str("."),
          pandoc.RawInline("field", string.format("{Seq equations|%q}", number.eq)),
          pandoc.Str(")"),
        }
        if elem.content[#elem.content].tag == "Str" then
          local identifier = elem.content[#elem.content].text
          if string.sub(identifier, 1, 5) == "{#eq:" and string.sub(identifier, -1) == "}" then
            identifier = string.sub(identifier, 3, -2)
            if numbers[identifier] then
              io.write(string.format("bookmark exist: %q\n", identifier))
            else
              numbers[identifier] = string.format("(%s.%s)", number[1], number.eq)
            end
            eqno = { pandoc.Span(eqno, { id=identifier }) }
          end
        end

        return { pandoc.Table({}, aligns, widths, { {}, {} }, { { { tcpr[1], pandoc.Div({ math }, { ["custom-style"]="Equation" }) }, { tcpr[2], pandoc.Div({ pandoc.Para(eqno) }, { ["custom-style"]="Equation Caption" }) } } }) }

      elseif elem.content[1].tag == "Image" and not elem.content[2] then
        local image = elem.content[1]

        if image.classes:includes("frame") then
          local class = table.concat(image.classes:filter(function (class) return string.find(class, "frame") == 1 end), " ")
          image.classes = image.classes:filter(function (class) return string.find(class, "frame") ~= 1 end)

          if #image.caption > 0 then
            number.fig = number.fig + 1
            local identifier = image.identifier
            image.identifier = ""
            if string.sub(identifier, 1, 4) == "fig:" then
              numbers[identifier] = string.format("图%s.%s", number[1], number.fig)
            end

            local caption = pandoc.Div({ pandoc.Para({ pandoc.Span(image.caption, { id=identifier }) }) }, { ["custom-style"]="Image Caption" })
            image.caption = {}
            return { pandoc.Div({ pandoc.Div({ elem }, { ["custom-style"]="Captioned Figure" }), caption }, { class=class }) }
          else
            return { pandoc.Div({ elem }, { class=class }) }
          end
        else
          if #image.caption > 0 then
            number.fig = number.fig + 1
            local identifier = image.identifier
            image.identifier = ""
            if string.sub(identifier, 1, 4) == "fig:" then
              numbers[identifier] = string.format("图%s.%s", number[1], number.fig)
            end

            image.caption = { pandoc.Span(image.caption, { id=identifier }) }
            return elem
          end
        end

      end
    end,
  }

  filters:insert{
    Cite = function (elem)
      local identifier = elem.citations[1].id
      if string.sub(identifier, 1, 3) == "eq:" then
        return {
          pandoc.Str("式"),
          pandoc.RawInline("field", string.format("{Ref %s|%s}", identifier, numbers[identifier] or "Error! Equation not defined")),
        }
      elseif string.sub(identifier, 1, 4) == "fig:" then
        return { pandoc.RawInline("field", string.format("{Ref %s \\r|%s}", identifier, numbers[identifier] or "Error! Figure not defined")) }
      elseif string.sub(identifier, 1, 4) == "tbl:" then
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

if configuration:includes("utility") then
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

    RawInline = function (elem)
      if elem.format:find("bookmark") then
        local text = ''
        if ("bookmarkStart"):find(elem.format) then text = text..string.format('<w:bookmarkStart w:id=%q w:name=%q/>', elem.text, elem.text) end
        if ("bookmarkEnd"):find(elem.format) then text = text..string.format('<w:bookmarkEnd w:id=%q/>', elem.text) end
        return pandoc.RawInline("openxml", text)

      elseif elem.format:find("field") then
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

if configuration:includes("bookmark_length") then
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
