-- Compact Markdown -> HTML class mapper for FlexUp legal styles.
-- Run with: pandoc -f markdown+fancy_lists -t html5 --standalone --lua-filter=filters/compact_to_html.lua in.md -o out.html
--
-- Mirrors filters/compact_to_docx.lua's recognition of the compact-markdown
-- conventions (Section/Article/Appendix headings, "1.1"/"a)"/"i." clause
-- lines, ::: Comments / ::: Definitions blocks) but targets CSS classes
-- consumed by styles/flexup_styles.css instead of Word custom styles.
--
-- Two differences from the DOCX filter are intentional, not oversights:
--  * Bullet lists stay native <ul> (not flattened into List/List 2/List 3
--    paragraphs) because the stylesheet numbers nesting depth via
--    `ul ul` / `ul ul ul` selectors, not per-item classes.
--  * Cross-reference links and `[]{#ref-...}` anchor spans pass through
--    untouched -- HTML anchors work natively, so none of the Word
--    bookmark-name hashing/REF-field machinery is needed.

local utils = pandoc.utils

local function trim(s)
  return (s:gsub("^%s+", ""):gsub("%s+$", ""))
end

local function styled_div(class_name, blocks, identifier)
  return pandoc.Div(blocks, pandoc.Attr(identifier or "", { class_name }, {}))
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

local function enum_name(v)
  if type(v) == "table" and v.t then
    return v.t
  end
  return tostring(v)
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

  local Token = str_text(out[idx])
  local class_alias = Token:match("^%{%.([%w%-%_]+)%}$")
  if class_alias then
    out:remove(idx)
    if #out > 0 and out[#out].t == "Space" then
      out:remove(#out)
    end
    return out, class_alias
  end

  local explicit_style = Token:match('^%{custom%-style="([^"]+)"%}$')
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

  local Token = str_text(inlines[idx])
  if Token:match("^%{%.([%w%-%_]+)%}$") then
    return true
  end
  if Token:match('^%{custom%-style="([^"]+)"%}$') then
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

-- Splits a paragraph containing "1.1 text / a) text / i. text"-style
-- compact clause lines (joined by soft breaks after the PowerShell
-- pre-pass dedents them) into individually classed divs.
local function convert_compact_clause_line(inlines, state)
  local cleaned, suffix_class = parse_inline_style_suffix(inlines)

  if suffix_class then
    return styled_div(suffix_class, { pandoc.Para(cleaned) }), nil
  end

  local anchors = extract_anchor_spans(cleaned)
  local without_anchors = remove_anchor_spans(cleaned)
  local text_md = inlines_to_markdown_line(without_anchors)

  if not state.in_appendix then
    local article2 = text_md:match("^%d+%.%d+%.?%s+(.+)$")
    if article2 then
      return styled_div("Article-2", { pandoc.Para(append_anchors(inlines_from_markdown(article2), anchors)) }), nil
    end
  end

  local alpha = text_md:match("^[a-zA-Z]%)%s+(.+)$")
  if alpha then
    local class_name = state.in_appendix and "Appendix-4" or "Article-3"
    return styled_div(class_name, { pandoc.Para(append_anchors(inlines_from_markdown(alpha), anchors)) }), nil
  end

  local roman = text_md:match("^[ivxlcdmIVXLCDM]+%.%s+(.+)$")
  if roman then
    local class_name = state.in_appendix and "Appendix-5" or "Article-4"
    return styled_div(class_name, { pandoc.Para(append_anchors(inlines_from_markdown(roman), anchors)) }), nil
  end

  return nil, cleaned
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

local function wrap_first_para_like(item_blocks, class_name)
  local out = pandoc.List:new()
  for _, b in ipairs(item_blocks) do
    out:insert(b)
  end

  for i, b in ipairs(out) do
    if b.t == "Div" and #b.attr.classes > 0 then
      return out
    end
    if b.t == "Para" then
      out[i] = styled_div(class_name, { b })
      return out
    end
    if b.t == "Plain" then
      out[i] = styled_div(class_name, { pandoc.Para(b.content) })
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
    return state.in_appendix and "Appendix-4" or "Article-3"
  end

  if list_style == "LowerRoman" then
    return state.in_appendix and "Appendix-5" or "Article-4"
  end

  if list_style == "Decimal" then
    return state.in_appendix and "Appendix-3" or "Article-2"
  end

  return nil
end

local ROMAN_VALUES = {
  ["I"] = 1, ["V"] = 5, ["X"] = 10, ["L"] = 50, ["C"] = 100, ["D"] = 500, ["M"] = 1000,
}

local function roman_to_arabic(roman)
  roman = roman:upper()
  local total = 0
  local prev = 0
  for i = #roman, 1, -1 do
    local v = ROMAN_VALUES[roman:sub(i, i)]
    if not v then
      return nil
    end
    if v < prev then
      total = total - v
    else
      total = total + v
      prev = v
    end
  end
  if total == 0 then
    return nil
  end
  return total
end

-- The compact-markdown convention cross-references appendices with plain
-- "(#appendix-1)"-style links rather than explicit [] {#...} anchors, so
-- Appendix/Article/Section headings get a matching id here -- pandoc would
-- otherwise never assign one, since these headings are rewritten into
-- plain Divs and lose the auto-identifier a native Header would get.
local function resolve_header_style(header, state)
  local text_md = inlines_to_markdown_line(header.content)

  if header.level == 1 then
    local appendix_num, appendix_title = text_md:match("^Appendix%s+(%d+)%.%s+(.+)$")
    if appendix_title then
      return "Appendix-1", inlines_from_markdown(appendix_title), "appendix-" .. appendix_num
    end
    return nil, nil, nil
  end

  if header.level == 2 then
    local section_roman, section_roman_title = text_md:match("^Section%s+([IVXLCDM]+)%.%s+(.+)$")
    if section_roman_title then
      local arabic = roman_to_arabic(section_roman)
      local identifier = arabic and ("section-" .. arabic) or nil
      return "Section", inlines_from_markdown(section_roman_title), identifier
    end
    local section_alpha_title = text_md:match("^Section%s+[A-Z]+%.%s+(.+)$")
    if section_alpha_title then
      return "Appendix-2", inlines_from_markdown(section_alpha_title), nil
    end
    return nil, nil, nil
  end

  if header.level == 3 then
    local article_num, article_title = text_md:match("^Article%s+(%d+)%.%s+(.+)$")
    if article_title then
      return "Article-1", inlines_from_markdown(article_title), "article-" .. article_num
    end

    if state.in_appendix then
      local appendix3_title = text_md:match("^%d+%.%s+(.+)$")
      if appendix3_title then
        return "Appendix-3", inlines_from_markdown(appendix3_title), nil
      end
    end
    return nil, nil, nil
  end

  return nil, nil, nil
end

local function is_definitions_div(div)
  local style = get_custom_style(div.attr)
  if style == "Definitions" then
    return true
  end
  for _, cls in ipairs(div.attr.classes or {}) do
    if cls == "Definitions" or cls == "definitions" then
      return true
    end
  end
  return false
end

-- A "::: Definitions" div holds a definition list that stands in for a
-- two-column term/definition table, exactly as in compact_to_docx.lua.
local function definitions_div_to_table(div, state, convert)
  local dl = nil
  for _, b in ipairs(div.content) do
    if b.t == "DefinitionList" then
      dl = b
      break
    end
  end
  if not dl or #dl.content == 0 then
    return nil
  end

  local function make_row(item)
    local term_blocks = convert({ pandoc.Plain(item[1]) }, state, { bullet_depth = 0 })

    local def_blocks = pandoc.List:new()
    for _, definition in ipairs(item[2] or {}) do
      for _, b in ipairs(definition) do
        def_blocks:insert(b)
      end
    end
    local converted_def = convert(def_blocks, state, { bullet_depth = 0 })
    if #converted_def == 0 then
      converted_def = pandoc.List:new({ pandoc.Para({}) })
    end

    return pandoc.Row({
      pandoc.Cell(term_blocks),
      pandoc.Cell(converted_def),
    })
  end

  local head_rows = pandoc.List:new()
  local body_rows = pandoc.List:new()
  for i, item in ipairs(dl.content) do
    if i == 1 then
      head_rows:insert(make_row(item))
    else
      body_rows:insert(make_row(item))
    end
  end

  local tbody = { attr = pandoc.Attr(), row_head_columns = 0, head = {}, body = body_rows }
  return pandoc.Table(
    { long = pandoc.Blocks({}) },
    {
      { pandoc.AlignDefault, 0.25 },
      { pandoc.AlignDefault, 0.75 },
    },
    pandoc.TableHead(head_rows),
    { tbody },
    pandoc.TableFoot({}),
    pandoc.Attr()
  )
end

local function convert_blocks(blocks, state, ctx)
  local out = pandoc.List:new()

  for _, block in ipairs(blocks) do
    if block.t == "Header" then
      local class_name, stripped, identifier = resolve_header_style(block, state)
      if class_name then
        append_blocks(out, styled_div(class_name, { pandoc.Para(stripped) }, identifier))
        if class_name == "Appendix-1" then
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

    elseif block.t == "Div" and is_definitions_div(block) then
      local tbl = definitions_div_to_table(block, state, convert_blocks)
      if tbl then
        out:insert(tbl)
      else
        append_blocks(out, convert_blocks(block.content, state, { bullet_depth = 0 }))
      end

    elseif block.t == "Div" then
      -- Fenced divs already carry their class (`::: Comments` -> class
      -- "Comments") which matches the stylesheet directly, so just recurse.
      local converted = convert_blocks(block.content, state, ctx)
      out:insert(pandoc.Div(converted, block.attr))

    elseif block.t == "OrderedList" then
      local item_class = ordered_list_style(block, state)
      local new_items = pandoc.List:new()
      for _, item in ipairs(block.content) do
        local converted_item = convert_blocks(item, state, {
          bullet_depth = ctx.bullet_depth,
        })
        if item_class then
          converted_item = wrap_first_para_like(converted_item, item_class)
          append_blocks(new_items, converted_item)
        else
          new_items:insert(converted_item)
        end
      end
      if item_class then
        append_blocks(out, new_items)
      else
        out:insert(pandoc.OrderedList(new_items, block.listAttributes))
      end

    elseif block.t == "Table" then
      -- Recurse into every cell so styled cell content (::: Comments,
      -- bullet lists, compact clause markers) is mapped to CSS classes,
      -- then rebuild the table so the conversion is captured reliably.
      local function rebuild_table_row(row)
        local cells = pandoc.List:new()
        for _, cell in ipairs(row.cells or {}) do
          local content = convert_blocks(cell.content or {}, state, { bullet_depth = 0 })
          cells:insert(pandoc.Cell(content, cell.alignment, cell.row_span, cell.col_span, cell.attr))
        end
        return pandoc.Row(cells, row.attr)
      end
      local function rebuild_table_rows(rows)
        local result = pandoc.List:new()
        if rows then
          for _, row in ipairs(rows) do result:insert(rebuild_table_row(row)) end
        end
        return result
      end

      local head_rows = rebuild_table_rows(block.head and block.head.rows)
      local body_rows = pandoc.List:new()
      if block.bodies then
        for _, tbody in ipairs(block.bodies) do
          for _, r in ipairs(rebuild_table_rows(tbody.head)) do body_rows:insert(r) end
          for _, r in ipairs(rebuild_table_rows(tbody.body)) do body_rows:insert(r) end
        end
      end
      local foot_rows = rebuild_table_rows(block.foot and block.foot.rows)

      local tbody = { attr = pandoc.Attr(), row_head_columns = 0, head = {}, body = body_rows }
      out:insert(pandoc.Table(
        block.caption,
        block.colspecs,
        pandoc.TableHead(head_rows),
        { tbody },
        pandoc.TableFoot(foot_rows),
        block.attr
      ))

    elseif block.t == "BulletList" then
      -- Left as a native list (not flattened): the stylesheet indents
      -- nesting depth via `ul ul` / `ul ul ul`, so recursing to strip any
      -- compact markers inside items is all that's needed here.
      local new_items = pandoc.List:new()
      for _, item in ipairs(block.content) do
        new_items:insert(convert_blocks(item, state, ctx))
      end
      out:insert(pandoc.BulletList(new_items))

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
