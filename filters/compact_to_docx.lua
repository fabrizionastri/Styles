-- Compact Markdown -> DOCX style mapper for FlexUp legal styles.
-- Run with: pandoc -f markdown+fancy_lists -t docx --reference-doc=styles/flexup_template.docx --lua-filter=filters/compact_to_docx.lua in.md -o out.docx

local utils = pandoc.utils

local ALIAS_TO_STYLE = {
  ["Comments"] = "Comments",
  ["Offset"] = "Offset",
  ["Published"] = "Published",
  ["List"] = "List",
  ["List-2"] = "List 2",
  ["List-3"] = "List 3",
  ["index-1"] = "index 1",
  ["index-2"] = "index 2",
  ["index-3"] = "index 3",
}

local BULLET_STYLE_BY_LEVEL = {
  [1] = "List",
  [2] = "List 2",
  [3] = "List 3",
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

local function alias_to_style(alias)
  if ALIAS_TO_STYLE[alias] then
    return ALIAS_TO_STYLE[alias]
  end
  if alias and alias ~= "" then
    return (alias:gsub("-", " "))
  end
  return nil
end

local function inlines_from_markdown(text)
  local doc = pandoc.read(text, "markdown-smart")
  if #doc.blocks == 0 then
    return pandoc.List:new()
  end
  local b = doc.blocks[1]
  if b.t == "Para" or b.t == "Plain" or b.t == "Header" then
    return b.content
  end
  return pandoc.List:new({ pandoc.Str(text) })
end

local function inlines_to_markdown_line(inlines)
  local tmp_doc = pandoc.Pandoc({ pandoc.Plain(inlines) }, pandoc.Meta({}))
  local txt = pandoc.write(tmp_doc, "markdown-smart")
  txt = txt:gsub("\r\n", "\n")
  txt = trim(txt:gsub("\n+$", ""))
  txt = txt:gsub("%s*\n%s*", " ")
  return txt
end

local function styled_div(style_name, blocks)
  return pandoc.Div(blocks, pandoc.Attr("", {}, { ["custom-style"] = style_name }))
end

local function copy_inlines(inlines)
  local out = pandoc.List:new()
  for _, inl in ipairs(inlines) do
    out:insert(inl)
  end
  return out
end

local function str_text(inl)
  if inl.text then
    return inl.text
  end
  if inl.c then
    return inl.c
  end
  return ""
end

local function parse_inline_style_suffix(inlines)
  local out = copy_inlines(inlines)
  local idx = #out
  while idx > 0 and (out[idx].t == "Space" or out[idx].t == "SoftBreak") do
    idx = idx - 1
  end
  if idx == 0 then
    return out, nil
  end
  if out[idx].t ~= "Str" then
    return out, nil
  end

  local token = str_text(out[idx])
  local class_alias = token:match("^%{%.([%w%-%_]+)%}$")
  if class_alias then
    out:remove(idx)
    if #out > 0 and out[#out].t == "Space" then
      out:remove(#out)
    end
    return out, alias_to_style(class_alias)
  end

  local explicit_style = token:match('^%{custom%-style="([^"]+)"%}$')
  if explicit_style then
    out:remove(idx)
    if #out > 0 and out[#out].t == "Space" then
      out:remove(#out)
    end
    return out, explicit_style
  end

  return out, nil
end

local function split_inlines_on_breaks(inlines)
  local lines = pandoc.List:new()
  local current = pandoc.List:new()

  for _, inl in ipairs(inlines) do
    if inl.t == "SoftBreak" or inl.t == "LineBreak" then
      lines:insert(current)
      current = pandoc.List:new()
    else
      current:insert(inl)
    end
  end

  lines:insert(current)
  return lines
end

local function line_has_compact_marker(inlines)
  local text = trim(utils.stringify(inlines))
  if text == "" then
    return false
  end
  return text:match("^%d+%.%d+%.?%s+.+$") ~= nil
    or text:match("^[a-zA-Z]%)%s+.+$") ~= nil
    or text:match("^[ivxlcdmIVXLCDM]+%.%s+.+$") ~= nil
end

local function line_has_style_suffix(inlines)
  local idx = #inlines
  while idx > 0 and (inlines[idx].t == "Space" or inlines[idx].t == "SoftBreak") do
    idx = idx - 1
  end
  if idx == 0 or inlines[idx].t ~= "Str" then
    return false
  end

  local token = str_text(inlines[idx])
  if token:match("^%{%.([%w%-%_]+)%}$") then
    return true
  end
  if token:match('^%{custom%-style="([^"]+)"%}$') then
    return true
  end
  return false
end

local function extract_anchor_spans(inlines)
  local anchors = pandoc.List:new()
  for _, inl in ipairs(inlines) do
    if inl.t == "Span" then
      local id = inl.attr and inl.attr.identifier or ""
      if id ~= "" and id:match("^ref%-") then
        anchors:insert(inl)
      end
    end
  end
  return anchors
end

local function append_anchors(inlines, anchors)
  if #anchors == 0 then
    return inlines
  end
  local out = pandoc.List:new()
  for _, inl in ipairs(inlines) do
    out:insert(inl)
  end
  for _, a in ipairs(anchors) do
    out:insert(pandoc.Space())
    out:insert(a)
  end
  return out
end

local function remove_anchor_spans(inlines)
  local out = pandoc.List:new()
  for _, inl in ipairs(inlines) do
    if inl.t == "Span" then
      local id = inl.attr and inl.attr.identifier or ""
      if id ~= "" and id:match("^ref%-") then
        goto continue
      end
    end
    out:insert(inl)
    ::continue::
  end
  return out
end

local function convert_compact_clause_line(inlines, state)
  local cleaned, suffix_style = parse_inline_style_suffix(inlines)

  if suffix_style then
    return styled_div(suffix_style, { pandoc.Para(cleaned) }), nil
  end

  local anchors = extract_anchor_spans(cleaned)
  local without_anchors = remove_anchor_spans(cleaned)
  local text_md = inlines_to_markdown_line(without_anchors)

  if not state.in_appendix then
    local article2 = text_md:match("^%d+%.%d+%.?%s+(.+)$")
    if article2 then
      return styled_div("Article 2", { pandoc.Para(append_anchors(inlines_from_markdown(article2), anchors)) }), nil
    end
  end

  local alpha = text_md:match("^[a-zA-Z]%)%s+(.+)$")
  if alpha then
    local style = state.in_appendix and "Appendix 4" or "Article 3"
    return styled_div(style, { pandoc.Para(append_anchors(inlines_from_markdown(alpha), anchors)) }), nil
  end

  local roman = text_md:match("^[ivxlcdmIVXLCDM]+%.%s+(.+)$")
  if roman then
    local style = state.in_appendix and "Appendix 5" or "Article 4"
    return styled_div(style, { pandoc.Para(append_anchors(inlines_from_markdown(roman), anchors)) }), nil
  end

  return nil, cleaned
end

local function is_br_inline(inl)
  if inl.t == "LineBreak" then
    return true
  end
  if inl.t == "RawInline" then
    local fmt = enum_name(inl.format)
    if fmt == "html" then
      local tag = trim((inl.text or ""):lower())
      return tag == "<br>" or tag == "<br/>" or tag == "<br />"
    end
  end
  return false
end

local function split_inlines_on_br(inlines)
  local segments = pandoc.List:new()
  local current = pandoc.List:new()

  for _, inl in ipairs(inlines) do
    if is_br_inline(inl) then
      segments:insert(current)
      current = pandoc.List:new()
    else
      current:insert(inl)
    end
  end

  segments:insert(current)
  return segments
end

local function split_cell_blocks_on_br(blocks)
  local out = pandoc.List:new()

  for _, b in ipairs(blocks) do
    if b.t == "Para" or b.t == "Plain" then
      local has_br = false
      for _, inl in ipairs(b.content) do
        if is_br_inline(inl) then
          has_br = true
          break
        end
      end

      if has_br then
        local segments = split_inlines_on_br(b.content)
        for _, seg in ipairs(segments) do
          while #seg > 0 and seg[1].t == "Space" do
            seg:remove(1)
          end
          while #seg > 0 and seg[#seg].t == "Space" do
            seg:remove(#seg)
          end
          if #seg > 0 then
            out:insert(pandoc.Para(seg))
          end
        end
      else
        out:insert(b)
      end
    else
      out:insert(b)
    end
  end

  return out
end

local function split_table_br(tbl)
  local function process_rows(rows)
    for _, row in ipairs(rows) do
      for _, cell in ipairs(row.cells or {}) do
        cell.content = split_cell_blocks_on_br(cell.content or {})
      end
    end
  end

  if tbl.head and tbl.head.rows then
    process_rows(tbl.head.rows)
  end

  if tbl.bodies then
    for _, body in ipairs(tbl.bodies) do
      if body.head then
        process_rows(body.head)
      end
      if body.body then
        process_rows(body.body)
      end
    end
  end

  return tbl
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

local function wrap_first_para_like(item_blocks, style_name)
  local out = pandoc.List:new()
  for _, b in ipairs(item_blocks) do
    out:insert(b)
  end

  for i, b in ipairs(out) do
    if b.t == "Div" and get_custom_style(b.attr) then
      return out
    end
    if b.t == "Para" then
      out[i] = styled_div(style_name, { b })
      return out
    end
    if b.t == "Plain" then
      out[i] = styled_div(style_name, { pandoc.Para(b.content) })
      return out
    end
    if b.t == "Header" then
      return out
    end
  end

  return out
end

local function ordered_list_style(ol, state)
  local list_style = enum_name(ol.style)

  if list_style == "LowerAlpha" then
    if state.in_appendix then
      return "Appendix 4"
    end
    return "Article 3"
  end

  if list_style == "LowerRoman" then
    if state.in_appendix then
      return "Appendix 5"
    end
    return "Article 4"
  end

  if list_style == "Decimal" then
    if state.in_appendix then
      return "Appendix 3"
    end
    return "Article 2"
  end

  return nil
end

local function resolve_header_style(header, state)
  local text_md = inlines_to_markdown_line(header.content)

  if header.level == 1 then
    local appendix_title = text_md:match("^Appendix%s+%d+%.%s+(.+)$")
    if appendix_title then
      return "Appendix 1", inlines_from_markdown(appendix_title)
    end
    return nil, nil
  end

  if header.level == 2 then
    local section_roman = text_md:match("^Section%s+[IVXLCDM]+%.%s+(.+)$")
    if section_roman then
      return "Section", inlines_from_markdown(section_roman)
    end
    local section_alpha = text_md:match("^Section%s+[A-Z]+%.%s+(.+)$")
    if section_alpha then
      return "Appendix 2", inlines_from_markdown(section_alpha)
    end
    return nil, nil
  end

  if header.level == 3 then
    local article_title = text_md:match("^Article%s+%d+%.%s+(.+)$")
    if article_title then
      return "Article 1", inlines_from_markdown(article_title)
    end

    if state.in_appendix then
      local appendix3_title = text_md:match("^%d+%.%s+(.+)$")
      if appendix3_title then
        return "Appendix 3", inlines_from_markdown(appendix3_title)
      end
    end
    return nil, nil
  end

  return nil, nil
end

local function convert_blocks(blocks, state, ctx)
  local out = pandoc.List:new()

  for _, block in ipairs(blocks) do
    if block.t == "Header" then
      local style_name, stripped = resolve_header_style(block, state)
      if style_name then
        append_blocks(out, styled_div(style_name, { pandoc.Para(stripped) }))
        if style_name == "Appendix 1" then
          state.in_appendix = true
        end
      else
        out:insert(block)
      end

    elseif block.t == "Para" or block.t == "Plain" then
      local lines = split_inlines_on_breaks(block.content)
      local has_compact_multiline = false
      if #lines > 1 then
        for _, line in ipairs(lines) do
          if line_has_compact_marker(line) or line_has_style_suffix(line) then
            has_compact_multiline = true
            break
          end
        end
      end

      if has_compact_multiline then
        for _, line in ipairs(lines) do
          if #line > 0 then
            local styled, fallback = convert_compact_clause_line(line, state)
            if styled then
              append_blocks(out, styled)
            else
              out:insert(pandoc.Para(fallback))
            end
          end
        end
      else
        local styled, fallback = convert_compact_clause_line(block.content, state)
        if styled then
          append_blocks(out, styled)
        else
          if block.t == "Para" then
            out:insert(pandoc.Para(fallback))
          else
            out:insert(pandoc.Plain(fallback))
          end
        end
      end

    elseif block.t == "Div" then
      local style_name = get_custom_style(block.attr)
      if not style_name then
        for _, cls in ipairs(block.attr.classes) do
          local mapped = alias_to_style(cls)
          if mapped then
            style_name = mapped
            break
          end
        end
      end

      local converted = convert_blocks(block.content, state, ctx)
      if style_name then
        append_blocks(out, styled_div(style_name, converted))
      else
        out:insert(pandoc.Div(converted, block.attr))
      end

    elseif block.t == "OrderedList" then
      local item_style = ordered_list_style(block, state)
      local new_items = pandoc.List:new()
      for _, item in ipairs(block.content) do
        local converted_item = convert_blocks(item, state, {
          bullet_depth = ctx.bullet_depth,
        })
        if item_style then
          converted_item = wrap_first_para_like(converted_item, item_style)
          append_blocks(new_items, converted_item)
        else
          new_items:insert(converted_item)
        end
      end
      if item_style then
        append_blocks(out, new_items)
      else
        out:insert(pandoc.OrderedList(new_items, block.listAttributes))
      end

    elseif block.t == "Table" then
      local tbl = split_table_br(block)
      if #tbl.colspecs == 2 then
        tbl.colspecs = {
          { pandoc.AlignDefault, 0.25 },
          { pandoc.AlignDefault, 0.75 },
        }
      end
      out:insert(tbl)

    elseif block.t == "BulletList" then
      local level = (ctx.bullet_depth or 0) + 1
      local bullet_style = BULLET_STYLE_BY_LEVEL[level] or "List 3"
      local flattened = pandoc.List:new()
      for _, item in ipairs(block.content) do
        local converted_item = convert_blocks(item, state, {
          bullet_depth = level,
        })
        converted_item = wrap_first_para_like(converted_item, bullet_style)
        append_blocks(flattened, converted_item)
      end

      append_blocks(out, flattened)

    else
      out:insert(block)
    end
  end

  return out
end

function Pandoc(doc)
  local state = { in_appendix = false }
  local blocks = convert_blocks(doc.blocks, state, { bullet_depth = 0 })
  return pandoc.Pandoc(blocks, doc.meta)
end
