# pandoc contract.md --lua-filter=legal.lua --reference-doc=reference.docx -o contract.docx
# pandoc -f markdown -t docx --reference-doc=contract.docx --lua-filter=filters/compact_to_docx.lua contract.slim.compact.md -o contract.out.docx
pandoc `
  -f markdown `
  -t docx  `
  --reference-doc=contract.docx  `
  --lua-filter=filters/compact_to_docx.lua  `
  test1.md  `
  -o test1.docx

