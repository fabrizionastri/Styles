**Developer Summary**
Goal: keep a reliable DOCX <-> Markdown contract workflow, while handling legacy Word documents by remapping styles in DOCX first.

**Current Toolchain**
- `d2m.ps1`: DOCX -> Markdown (uses `filters/docx_to_compact.lua`)
- `m2d.ps1`: Markdown -> DOCX (uses `filters/compact_to_docx.lua` and `styles.docx`)
- `commands.ps1`: loads `d2m` and `m2d` helper commands in PowerShell
- `install_commands.ps1`: adds commands loader to your PowerShell profile
- `remap_legacy_contracts.py`: remaps legacy Word heading styles before conversion

**Legacy Remap (DOCX First)**
Use the Python script before `d2m` when a legacy contract does not follow current styles.

Style mapping implemented in `remap_legacy_contracts.py`:
- `Title` -> `Heading 1`
- `Heading 1` -> `Article 1`
- `Heading 2` -> `Article 2`
- `Heading 3` -> `Article 3`
- `Heading 4` -> `Article 4`
- `Heading 6` -> `Heading 4`

Default output naming:
- Input `legacy.docx` -> output `legacy_remapped.docx`
- If output is provided without extension, `.docx` is appended.
- If input is provided without extension, `.docx` is inferred.

Run examples:
```powershell
# In activated virtual environment
python .\remap_legacy_contracts.py legacy
python .\remap_legacy_contracts.py legacy legacy_clean

# Alternative without activating venv
uv run python .\remap_legacy_contracts.py legacy
```

Expected output message:
```text
Success! Saved to legacy_remapped.docx
```

**Conversion Commands**
DOCX -> Markdown:
```powershell
powershell -ExecutionPolicy Bypass -File .\d2m.ps1 "SaaS contract.docx"
powershell -ExecutionPolicy Bypass -File .\d2m.ps1 "SaaS contract.docx" "plop.md"
```

Markdown -> DOCX:
```powershell
powershell -ExecutionPolicy Bypass -File .\m2d.ps1 "SaaS contract.md"
powershell -ExecutionPolicy Bypass -File .\m2d.ps1 "SaaS contract.md" "plop.docx"
```

Extension inference:
- `d2m` input defaults to `.docx`, output defaults to `.md`
- `m2d` input defaults to `.md`, output defaults to `.docx`
- Example: `d2m "style" "plop"` -> `style.docx` to `plop.md`
- Example: `m2d "plop" "zut"` -> `plop.md` to `zut.docx`

**Optional Shell Helpers**
Load functions in the current session:
```powershell
. .\commands.ps1
d2m "SaaS contract.docx"
m2d "SaaS contract.md"
```

Install for every new PowerShell session:
```powershell
powershell -ExecutionPolicy Bypass -File .\install_commands.ps1
```

**Recommended Legacy Workflow**
1. Remap legacy DOCX styles:
```powershell
python .\remap_legacy_contracts.py "legacy_contract.docx"
```
2. Convert remapped DOCX to Markdown:
```powershell
d2m "legacy_contract_remapped.docx"
```
3. Edit Markdown.
4. Convert back to DOCX:
```powershell
m2d "legacy_contract_remapped.md"
```

**Notes**
- The legacy mapping is now handled in Python (`python-docx`), not in Lua filters.
- Keep `styles.docx` aligned with your target Word style definitions, since it is used as reference during `m2d`.
