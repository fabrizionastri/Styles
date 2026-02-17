pandoc -f docx+styles -t markdown --lua-filter=filters/docx_to_compact.lua styles.docx -o example.md
