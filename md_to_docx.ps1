pandoc -f markdown+fancy_lists+lists_without_preceding_blankline -t docx --reference-doc=styles.docx --lua-filter=filters/compact_to_docx.lua example.md -o example.docx
