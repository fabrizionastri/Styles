-- Legacy Word heading mapper.
-- Intended for old contracts that used Heading 1..4/6 instead of the new legal styles.
--
-- Mapping:
--   Heading 1 -> Article 1  (### Article N. <text>)
--   Heading 2 -> Article 2  (<N.M. <text>> paragraph line)
--   Heading 3 -> Article 3  (<a) <text>> paragraph line)
--   Heading 4 -> Article 4  (<i. <text>> paragraph line)
--   Heading 6 -> Heading 4  (#### <text>)

local utils = pandoc.utils

local state = {
  article1 = 0,
  article2 = 0,
  article3 = 0,
  article4 = 0,
}

local function to_alpha(n)
  local out = ""
  local x = n
  while x > 0 do
    local r = (x - 1) % 26
    out = string.char(string.byte("a") + r) .. out
    x = math.floor((x - 1) / 26)
  end
  return out
end

local function to_roman(n)
  local vals = {
    {1000, "m"}, {900, "cm"}, {500, "d"}, {400, "cd"}, {100, "c"},
    {90, "xc"}, {50, "l"}, {40, "xl"}, {10, "x"}, {9, "ix"},
    {5, "v"}, {4, "iv"}, {1, "i"},
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

function Header(h)
  local text = utils.stringify(h.content)

  if h.level == 1 then
    state.article1 = state.article1 + 1
    state.article2 = 0
    state.article3 = 0
    state.article4 = 0
    return pandoc.Header(3, inlines_from_markdown("Article " .. state.article1 .. ". " .. text))
  end

  if h.level == 2 then
    if state.article1 == 0 then
      state.article1 = 1
    end
    state.article2 = state.article2 + 1
    state.article3 = 0
    state.article4 = 0
    return pandoc.Plain(inlines_from_markdown(state.article1 .. "." .. state.article2 .. ". " .. text))
  end

  if h.level == 3 then
    if state.article1 == 0 then
      state.article1 = 1
    end
    if state.article2 == 0 then
      state.article2 = 1
    end
    state.article3 = state.article3 + 1
    state.article4 = 0
    return pandoc.Plain(inlines_from_markdown(to_alpha(state.article3) .. ") " .. text))
  end

  if h.level == 4 then
    if state.article1 == 0 then
      state.article1 = 1
    end
    if state.article2 == 0 then
      state.article2 = 1
    end
    if state.article3 == 0 then
      state.article3 = 1
    end
    state.article4 = state.article4 + 1
    return pandoc.Plain(inlines_from_markdown(to_roman(state.article4) .. ". " .. text))
  end

  if h.level == 6 then
    return pandoc.Header(4, h.content)
  end

  return h
end

