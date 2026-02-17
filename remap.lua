function Header(el)
  -- Map Heading 1-4 to Article 1-4
  -- We turn them into Paragraphs with a custom style attribute
  if el.level >= 1 and el.level <= 4 then
    local style_name = "Article " .. el.level
    return pandoc.Para(el.content, {['custom-style'] = style_name})
  end
  
  -- Map Heading 6 to Heading 4
  if el.level == 6 then
    el.level = 4
    return el
  end

end

function Para(el)
  -- Map Title style to Heading 1
  if el.attributes['custom-style'] == 'Title' then
    return pandoc.Header(1, el.content)
  end
end
