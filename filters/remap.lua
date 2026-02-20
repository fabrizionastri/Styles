--[[ Run with:
pandoc -f docx+styles legacy.docx --lua-filter=filters/remap.lua --reference-doc=styles/contract_template.docx --no-highlight -o output.docx
--]]


local STYLE_MAP = {
  ["Title"] = "Heading 1",
  ["Heading 1"] = "Article 1",
  ["Heading 2"] = "Article 2",
  ["Heading 3"] = "Article 3",
  ["Heading 4"] = "Article 4",
  ["Heading 6"] = "Heading 4",
  ["PublishedOn"] = "Published",
}

local function get_custom_style(attr)
  if not attr or not attr.attributes then
    return nil
  end
  return attr.attributes["custom-style"]
end

local function attr_with_custom_style(attr, style_name)
  local identifier = ""
  local classes = {}
  local attributes = {}

  if attr then
    identifier = attr.identifier or ""

    if attr.classes then
      for _, class_name in ipairs(attr.classes) do
        classes[#classes + 1] = class_name
      end
    end

    if attr.attributes then
      for key, value in pairs(attr.attributes) do
        attributes[key] = value
      end
    end
  end

  attributes["custom-style"] = style_name
  return pandoc.Attr(identifier, classes, attributes)
end

function Div(el)
  local source_style = get_custom_style(el.attr)
  local target_style = STYLE_MAP[source_style]
  if not target_style then
    return nil
  end
  return pandoc.Div(el.content, attr_with_custom_style(el.attr, target_style))
end

function Header(el)
  if el.level >= 1 and el.level <= 4 then
    local target_style = "Article " .. tostring(el.level)
    return pandoc.Div(
      { pandoc.Para(el.content) },
      pandoc.Attr("", {}, { ["custom-style"] = target_style })
    )
  end

  if el.level == 6 then
    return pandoc.Div(
      { pandoc.Para(el.content) },
      pandoc.Attr("", {}, { ["custom-style"] = "Heading 4" })
    )
  end

  return nil
end

local function meta_title_to_inlines(meta_title)
  if not meta_title then
    return nil
  end

  local value_type = pandoc.utils.type(meta_title)

  if value_type == "Inlines" then
    return meta_title
  end

  if value_type == "MetaInlines" then
    return meta_title.c
  end

  if value_type == "MetaString" then
    return { pandoc.Str(tostring(meta_title)) }
  end

  if value_type == "MetaBlocks" then
    return pandoc.utils.blocks_to_inlines(meta_title)
  end

  if type(meta_title) == "table" and meta_title.t == "MetaInlines" then
    return meta_title.c
  end

  if type(meta_title) == "table" and meta_title.t == "MetaString" then
    return { pandoc.Str(meta_title.c) }
  end

  if type(meta_title) == "table" and meta_title.t == "MetaBlocks" then
    return pandoc.utils.blocks_to_inlines(meta_title.c)
  end

  return nil
end

function Pandoc(doc)
  local title_inlines = meta_title_to_inlines(doc.meta and doc.meta.title)
  if not title_inlines or #title_inlines == 0 then
    return doc
  end

  local title_div = pandoc.Div(
    { pandoc.Para(title_inlines) },
    pandoc.Attr("", {}, { ["custom-style"] = "Heading 1" })
  )

  local updated_blocks = { title_div }
  for _, block in ipairs(doc.blocks) do
    updated_blocks[#updated_blocks + 1] = block
  end

  doc.blocks = updated_blocks
  if doc.meta then
    doc.meta.title = nil
  end
  return doc
end
