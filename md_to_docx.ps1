pandoc -f markdown+fancy_lists -t docx --reference-doc=styles.docx --lua-filter=filters/compact_to_docx.lua example.md -o example.docx
