-- DOCX -> compact Markdown normalizer for FlexUp legal styles.
-- Run with: pandoc -f docx+styles -t markdown --lua-filter=filters/docx_to_compact.lua in.docx -o out.md

local utils = pandoc.utils

local HEADING_STYLE_MAP = {
  ["Section"] = { level = 2, prefix = "Section ", counter = "roman_upper" },
  ["Article 1"] = { level = 3, prefix = "Article ", counter = "decimal" },
  ["Appendix 1"] = { level = 1, prefix = "Appendix ", counter = "decimal" },
  ["Appendix 2"] = { level = 2, prefix = "Section ", counter = "alpha_upper" },
  ["Appendix 3"] = { level = 3, prefix = "", counter = "decimal" },
}

local UNWRAP_STYLE_SET = {
  ["Article 2"] = true,
  ["Article 3"] = true,
  ["Article 4"] = true,
  ["Appendix 4"] = true,
  ["Appendix 5"] = true,
  ["List"] = true,
  ["List 2"] = true,
  ["List 3"] = true,
  ["endnote text"] = true,
  ["footnote text"] = true,
}

local function trim(s)
  return (s:gsub("^%s+", ""):gsub("%s+$", ""))
end

local function enum_name(v)
  if type(v) == "table" and v.t then
    return v.t
  end
  return tostring(v)
end

local function to_alpha(n, uppercase)
  local out = ""
  while n > 0 do
    local r = (n - 1) % 26
    out = string.char(string.byte("A") + r) .. out
    n = math.floor((n - 1) / 26)
  end
  if uppercase then
    return out
  end
  return out:lower()
end

local function to_roman(n, uppercase)
  local vals = {
    {1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"}, {100, "C"},
    {90, "XC"}, {50, "L"}, {40, "XL"}, {10, "X"}, {9, "IX"},
    {5, "V"}, {4, "IV"}, {1, "I"},
  }
  local out = ""
  local x = n
  for _, pair in ipairs(vals) do
    local v, glyph = pair[1], pair[2]
    while x >= v do
      out = out .. glyph
      x = x - v
    end
  end
  if uppercase then
    return out
  end
  return out:lower()
end

local function format_counter(counter_kind, n)
  if counter_kind == "decimal" then
    return tostring(n)
  end
  if counter_kind == "alpha_upper" then
    return to_alpha(n, true)
  end
  if counter_kind == "roman_upper" then
    return to_roman(n, true)
  end
  return tostring(n)
end

local function get_custom_style(attr)
  if not attr then
    return nil
  end
  if attr.attributes and attr.attributes["custom-style"] then
    return attr.attributes["custom-style"]
  end
  local kv = attr[3]
  if kv then
    for _, pair in ipairs(kv) do
      if pair[1] == "custom-style" then
        return pair[2]
      end
    end
  end
  return nil
end

local function style_to_alias(style)
  return (style:gsub("%s+", "-"))
end

local function append_suffix_class(inlines, alias)
  local out = pandoc.List:new()
  for _, inl in ipairs(inlines) do
    out:insert(inl)
  end
  if #out > 0 then
    out:insert(pandoc.Space())
  end
  out:insert(pandoc.Str("{." .. alias .. "}"))
  return out
end

local function inlines_from_markdown(text)
  local doc = pandoc.read(text, "markdown")
  if #doc.blocks == 0 then
    return pandoc.List:new()
  end
  local b = doc.blocks[1]
  if b.t == "Para" or b.t == "Plain" or b.t == "Header" then
    return b.content
  end
  return pandoc.List:new({ pandoc.Str(text) })
end

local function clean_inlines(inlines, state, convert_blocks)
  local out = pandoc.List:new()
  for _, inl in ipairs(inlines) do
    if inl.t == "Span" then
      local span_style = get_custom_style(inl.attr)
      local inner = clean_inlines(inl.content, state, convert_blocks)
      if span_style and span_style:lower():find("note reference", 1, true) then
        for _, s in ipairs(inner) do
          out:insert(s)
        end
      else
        out:insert(pandoc.Span(inner))
      end
    elseif inl.t == "Note" then
      local note_blocks = convert_blocks(inl.content, state)
      out:insert(pandoc.Note(note_blocks))
    else
      out:insert(inl)
    end
  end
  return out
end

local function extract_primary_inlines(blocks)
  for _, b in ipairs(blocks) do
    if b.t == "Para" or b.t == "Plain" then
      return b.content
    end
  end
  return pandoc.List:new()
end

local function append_blocks(dst, src)
  if src == nil then
    return
  end
  if src.t ~= nil then
    dst:insert(src)
    return
  end
  for _, b in ipairs(src) do
    dst:insert(b)
  end
end

local function convert_blocks(blocks, state)
  local function convert_div(div)
    local style = get_custom_style(div.attr)
    local converted_content = convert_blocks(div.content, state)

    if not style then
      return pandoc.Div(converted_content, div.attr)
    end

    local lower_style = style:lower()
    if lower_style:find("toc", 1, true) then
      return pandoc.List:new()
    end

    if style == "Comments" or style == "Published" then
      local inlines = extract_primary_inlines(converted_content)
      return pandoc.Para(append_suffix_class(inlines, style_to_alias(style)))
    end

    if UNWRAP_STYLE_SET[style] then
      return converted_content
    end

    if #converted_content == 1 and (converted_content[1].t == "Para" or converted_content[1].t == "Plain") then
      local alias = style_to_alias(style)
      return pandoc.Para(append_suffix_class(converted_content[1].content, alias))
    end

    local alias = style_to_alias(style)
    return pandoc.Div(converted_content, pandoc.Attr("", {alias}, {}))
  end

  local function convert_ordered_list(ol)
    local attrs = ol.listAttributes
    local start = ol.start or 1

    local all_heading_items = true
    for _, item in ipairs(ol.content) do
      local first = item[1]
      if not first or first.t ~= "Div" then
        all_heading_items = false
        break
      end
      local style = get_custom_style(first.attr)
      if not HEADING_STYLE_MAP[style] then
        all_heading_items = false
        break
      end
    end

    if all_heading_items then
      local out = pandoc.List:new()
      for i, item in ipairs(ol.content) do
        local first = item[1]
        local style = get_custom_style(first.attr)
        local spec = HEADING_STYLE_MAP[style]
        local n = start + (i - 1)

        local first_blocks = convert_blocks(first.content, state)
        local heading_inlines = extract_primary_inlines(first_blocks)
        local marker = format_counter(spec.counter, n)
        local prefix = spec.prefix .. marker .. "."
        local prefixed = inlines_from_markdown(prefix)
        if #heading_inlines > 0 then
          prefixed:insert(pandoc.Space())
        end
        for _, inl in ipairs(heading_inlines) do
          prefixed:insert(inl)
        end

        out:insert(pandoc.Header(spec.level, prefixed))

        if style == "Appendix 1" then
          state.in_appendix = true
        end

        local remainder = pandoc.List:new()
        if #first_blocks > 1 then
          for j = 2, #first_blocks do
            remainder:insert(first_blocks[j])
          end
        end
        for j = 2, #item do
          remainder:insert(item[j])
        end
        append_blocks(out, convert_blocks(remainder, state))
      end
      return out
    end

    local new_items = pandoc.List:new()
    for _, item in ipairs(ol.content) do
      new_items:insert(convert_blocks(item, state))
    end
    return pandoc.OrderedList(new_items, attrs)
  end

  local function convert_bullet_list(bl)
    local new_items = pandoc.List:new()
    for _, item in ipairs(bl.content) do
      new_items:insert(convert_blocks(item, state))
    end
    return pandoc.BulletList(new_items)
  end

  local out = pandoc.List:new()

  for _, block in ipairs(blocks) do
    if block.t == "RawBlock" and enum_name(block.format) == "html" then
      local raw = trim(block.text or block.c or "")
      if raw == "<!-- -->" then
        goto continue
      end
    end

    if block.t == "Header" then
      local txt = trim(utils.stringify(block.content))
      if block.level == 1 and txt:match("^Appendix%s+%d+%.") then
        state.in_appendix = true
      end
      block.content = clean_inlines(block.content, state, convert_blocks)
      out:insert(block)
    elseif block.t == "Para" then
      block.content = clean_inlines(block.content, state, convert_blocks)
      out:insert(block)
    elseif block.t == "Plain" then
      block.content = clean_inlines(block.content, state, convert_blocks)
      out:insert(block)
    elseif block.t == "Div" then
      append_blocks(out, convert_div(block))
    elseif block.t == "OrderedList" then
      append_blocks(out, convert_ordered_list(block))
    elseif block.t == "BulletList" then
      append_blocks(out, convert_bullet_list(block))
    else
      out:insert(block)
    end

    ::continue::
  end

  return out
end

function Pandoc(doc)
  local state = { in_appendix = false }
  local blocks = convert_blocks(doc.blocks, state)
  return pandoc.Pandoc(blocks, doc.meta)
end
