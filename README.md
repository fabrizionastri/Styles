# FlexUp Contract Toolbox

## Why this toolbox?

We currently manage a dozen contract templates, and we are already using AI to update them and to ensure coherence across different templates. In the near future, we will be creating and managing hundreds of templates in different languages. To do this efficiently at scale, we need a better way to work with AI and to compare different versions of our documents.

AI tools work much better with Markdown (.md) files than with Word (.docx) files. Markdown files are also much easier to compare side-by-side using tools like [VS Code](https://code.visualstudio.com/download) or [Beyond Compare](https://www.scootersoftware.com/download), which makes reviewing changes across templates far more practical.

This toolbox lets you convert freely between Word and Markdown:

- **Word to Markdown** -- so you can feed the document to AI tools or compare versions
- **Markdown to Word** -- so you can get back a properly formatted Word document

To make these conversions reliable, we have redesigned the full set of Word styles. This means that for our existing (legacy) contracts, we first need to convert them to the new Word style set, verify the result in Word, and only then convert them to Markdown.

---

## Quick start

### Step 1. Install Pandoc (one-time)

Pandoc is the engine that converts between Word and Markdown. You only need to install it once.

1. Open your web browser and go to: https://pandoc.org/installing.html
2. Under **Windows**, click the download link for the latest `.msi` installer
3. Run the downloaded file and follow the installation wizard (accept all defaults)
4. When it finishes, close any open PowerShell windows (the new program will only be available after you reopen them)

### Step 2. Install the commands (one-time)

1. Open PowerShell (press the Windows key, type `PowerShell`, and click on it)
2. Navigate to this folder. For example, if this folder is on your Desktop:
   ```
   cd "$HOME\Desktop\Styles"
   ```
   Adjust the path to wherever this folder actually is on your computer.
3. Run the installer:
   ```
   powershell -ExecutionPolicy Bypass -File .\commands\install_commands.ps1
   ```
   You should see a message saying the commands were added to your profile.
4. **Close PowerShell and open it again.** This is necessary for the new commands to become available.

### Step 3. Use the commands

From now on, every time you open PowerShell, the three commands are ready to use. Just navigate to the folder that contains your contract files and run them.

**Convert Word to Markdown:**
```
d2m "My Contract.docx"
```
This creates `My Contract.md` in the same folder.

**Convert Markdown to Word:**
```
m2d "My Contract.md"
```
This creates `My Contract.docx` in the same folder.

**Convert a legacy Word document to the new styles:**
```
ld2d "Old Contract.docx"
```
This creates `Old Contract_remapped.docx` in the same folder. Open this file in Word and verify that it looks correct before proceeding.

You can also specify a different output name:
```
d2m "My Contract.docx" "output.md"
m2d "My Contract.md" "output.docx"
ld2d "Old Contract.docx" "New Contract"
```

You don't need to type the file extension -- the commands will add `.docx` or `.md` automatically if you leave it out.

---

## Recommended workflow

### Working on a document

When you are doing heavy editing on a document (drafting, reviewing, formatting), **work in Word**. Word gives you the best editing experience for legal text. Once you finish your manual review, convert the final version to Markdown:

```
d2m "My Contract.docx"
```

### Working with AI

When you need AI to review, modify, or compare your documents, **provide the Markdown file**. AI tools handle Markdown much more effectively than Word files.

After the AI produces an updated Markdown file, convert it back to Word:

```
m2d "My Contract.md"
```

### Comparing documents

To compare two versions of a document, convert both to Markdown and use a comparison tool such as VS Code (free) or Beyond Compare. Markdown files make differences between versions immediately visible, which is very useful for reviewing changes across multiple templates.

---

## Handling legacy contracts

Existing contracts that were created before the new style set need to be converted before they can be used with this toolbox.

1. **Convert to the new styles:**
   ```
   ld2d "legacy_contract.docx"
   ```
   This produces `legacy_contract_remapped.docx`.

2. **Open the remapped file in Word** and verify that all headings, articles, and formatting look correct. Fix anything that looks off and save.

3. **Convert to Markdown:**
   ```
   d2m "legacy_contract_remapped.docx"
   ```

4. From this point forward, you can edit the Markdown or the Word version and convert between them freely using `d2m` and `m2d`.

---

## Troubleshooting

**"pandoc is not installed or not in PATH"**
Pandoc is not installed, or PowerShell was not restarted after installation. Install Pandoc (see Step 1 above) and open a new PowerShell window.

**"d2m is not recognized as a command"**
The commands are not loaded. Either run `install_commands.ps1` again (see Step 2 above), or close and reopen PowerShell.

**"Input file not found"**
Check that the file name is correct and that you are in the right folder. Use `dir` to list the files in the current folder. If the file name contains spaces, make sure to wrap it in quotes: `d2m "My Contract.docx"`.

---

---

# Technical Reference

Everything below is detailed technical documentation for developers and advanced users.

## Toolchain overview

| Script                               | Purpose                                            | Key dependencies                                               |
| ------------------------------------ | -------------------------------------------------- | -------------------------------------------------------------- |
| `commands/d2m.ps1`                   | DOCX to Markdown                                   | `filters/docx_to_compact.lua`                                  |
| `commands/m2d.ps1`                   | Markdown to DOCX                                   | `filters/compact_to_docx.lua`, `styles/contract_template.docx` |
| `commands/ld2d.ps1`                  | Legacy DOCX to remapped DOCX                       | `filters/remap.lua`, `styles/contract_template.docx`           |
| `commands/commands.ps1`              | Loads `d2m`, `m2d`, `ld2d` as PowerShell functions | --                                                             |
| `commands/install_commands.ps1`      | Adds commands loader to your PowerShell profile    | --                                                             |
| `commands/remap_legacy_contracts.py` | Python alternative for legacy style remapping      | `python-docx`                                                  |

## Conversion commands (direct invocation)

If you prefer to run the scripts directly instead of using the shell helpers:

**DOCX to Markdown:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\d2m.ps1 "SaaS contract.docx"
powershell -ExecutionPolicy Bypass -File .\commands\d2m.ps1 "SaaS contract.docx" "output.md"
```

**Markdown to DOCX:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\m2d.ps1 "SaaS contract.md"
powershell -ExecutionPolicy Bypass -File .\commands\m2d.ps1 "SaaS contract.md" "output.docx"
```

**Legacy DOCX to remapped DOCX:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\ld2d.ps1 "legacy.docx"
powershell -ExecutionPolicy Bypass -File .\commands\ld2d.ps1 "legacy.docx" "output"
```

## Extension inference

All commands automatically add file extensions when omitted:

| Command | Input default | Output default   |
| ------- | ------------- | ---------------- |
| `d2m`   | `.docx`       | `.md`            |
| `m2d`   | `.md`         | `.docx`          |
| `ld2d`  | `.docx`       | `_remapped.docx` |

Examples:
- `d2m "contract"` reads `contract.docx`, writes `contract.md`
- `m2d "contract" "final"` reads `contract.md`, writes `final.docx`
- `ld2d "legacy"` reads `legacy.docx`, writes `legacy_remapped.docx`
- `ld2d "legacy.docx" "clean"` reads `legacy.docx`, writes `clean.docx`

## Legacy style mapping

The legacy remap (both the Lua filter in `ld2d` and the Python script) applies the following style conversions:

| Legacy style | New style |
| ------------ | --------- |
| Title        | Heading 1 |
| Heading 1    | Article 1 |
| Heading 2    | Article 2 |
| Heading 3    | Article 3 |
| Heading 4    | Article 4 |
| Heading 6    | Heading 4 |

The Lua filter (`remap.lua`) additionally converts native Markdown headings (H1-H4) to their corresponding Article styles, and H6 to Heading 4.

## Python legacy remap (alternative)

An alternative to `ld2d` is the Python script, which remaps styles directly in the DOCX file without going through Pandoc:

```powershell
# With activated virtual environment
python .\commands\remap_legacy_contracts.py legacy
python .\commands\remap_legacy_contracts.py legacy legacy_clean

# Without activating the virtual environment
uv run python .\commands\remap_legacy_contracts.py legacy
```

This requires Python 3.11+ and the `python-docx` package (managed via `pyproject.toml` and `uv`).

## Reference document

The file `styles/contract_template.docx` is the Word reference template used by `m2d` and `ld2d` to produce correctly styled output. Keep this file aligned with your target Word style definitions. See the [style guide](styles/style_guide.md) and [style matrix](styles/style_matrix.md) for full style specifications.

## Related documentation

- [styles/style_guide.md](styles/style_guide.md) -- full specification of every paragraph style, numbering scheme, and Markdown mapping
- [styles/style_matrix.md](styles/style_matrix.md) -- compact reference table of all styles
- [styles/contract_template.md](styles/contract_template.md) -- example Markdown contract showing all styles in use
