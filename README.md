# FlexUp Contract Commons

Open contract templates and conversion tools for working between Word (`.docx`) and Markdown (`.md`).

## Why this project?

This repository has two goals:

1. Build and maintain an open contract library that anyone can use and improve.
2. Provide reliable conversion tools so contracts can move between Word and Markdown without losing structure and style.

Markdown works well for AI-assisted drafting, versioning, and diff review, while Word remains the most practical format for legal editing and final review. This toolbox keeps both workflows connected.

## Licence and use

The contract library and related materials are published under the **FlexUp Licence**. See `licence.md` for the full terms.

## Contributions welcome

This is a public repository, and contributions are encouraged.

You can contribute by:

- adding new contract templates
- improving existing clauses and template quality
- adding translations / jurisdiction variants
- improving conversion reliability (`docx <-> md`)
- improving filters, scripts, tests, and documentation

## What this toolbox does

It converts freely between Word and Markdown:

- **Word to Markdown** -- so you can feed the document to AI tools or compare versions
- **Markdown to Word** -- so you can get back a properly formatted Word document

For legacy contracts, first remap old Word styles to the new style set, verify in Word, then convert to Markdown.



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

**Convert Markdown to HTML:**
```
m2h "My Contract.md"
```
This creates `My Contract.html` in the same folder — a single self-contained file with the FlexUp styling embedded, ready to open in a browser or attach in an email.

**Convert a legacy Word document to the new styles:**
```
ld2d "Old Contract.docx"
```
This creates `Old Contract_remapped.docx` in the same folder. Open this file in Word and verify that it looks correct before proceeding.

**Convert all Markdown files in a folder to Word:**
```
m2d-all
m2d-all "C:\Contracts\Templates"
```
With no argument, converts every `.md` file in the current folder. With a folder path, converts every `.md` file in that folder. Output files are always created alongside the source files.

**Convert all Word files in a folder to Markdown:**
```
d2m-all
d2m-all "C:\Contracts\Templates"
```
Same as above, but converts every `.docx` file to `.md`.

**Convert all Markdown files in a folder to HTML:**
```
m2h-all
m2h-all "C:\Contracts\Templates"
```
Same idea, but converts every `.md` file to `.html`.

If an output file already exists, the commands add a counter to avoid overwriting: `My Contract_1.docx`, `My Contract_2.docx`, etc.

You can also specify where single-file commands should write their output:
```
d2m "My Contract.docx" "output.md"
m2d "My Contract.md" "output.docx"
m2h "My Contract.md" "output.html"
ld2d "Old Contract.docx" "New Contract"
d2m "My Contract.docx" "."
d2m "My Contract.docx" ".\plop"
```

You don't need to type the file extension -- the commands will add `.docx` or `.md` automatically if you leave it out.

When you do not provide a second argument, the output is created in the same folder as the input file. When you do provide one, relative paths are resolved from your current PowerShell folder:

- `.` means "write here" and keeps the input file name, for example `My Contract.md`
- `.\plop` means "write here as `plop.md`" for `d2m`, or `plop.docx` for `m2d`
- an existing folder path means "write into that folder" and keep the input file name



### Step 4. Set up the VS Code shortcut (one-time, optional)

If you draft documents in Markdown and want to paste them into Outlook with formatting intact, install this keyboard shortcut. It converts the active `.md` file to HTML and places it on the Windows clipboard in a format that Outlook recognises as rich text — so paragraphs stay as paragraphs instead of turning into bullet points.

**Requirements:** VS Code, plus Python and Pandoc (both covered by Steps 1 and 2 above).

1. **Copy the script.** Create the folder `%APPDATA%\Code\User\scripts\` if it does not already exist, then copy `commands/md2clip.py` from this repository into it. The destination path will look like:
   ```
   C:\Users\<your-name>\AppData\Roaming\Code\User\scripts\md2clip.py
   ```
   Tip: press Win+R, type `%APPDATA%`, and press Enter to open the folder directly in Windows Explorer.

2. **Create the VS Code task.** Create the file `%APPDATA%\Code\User\tasks.json` with the content below. If the file already exists, add the task object to the existing `"tasks"` array instead of replacing the whole file.
   ```json
   {
     "version": "2.0.0",
     "tasks": [
       {
         "label": "Copy MD as HTML",
         "type": "process",
         "command": "python",
         "args": [
           "${env:APPDATA}/Code/User/scripts/md2clip.py",
           "${file}"
         ],
         "problemMatcher": [],
         "presentation": {
           "reveal": "silent",
           "revealProblems": "onProblem",
           "close": true,
           "panel": "dedicated"
         }
       }
     ]
   }
   ```

3. **Add the keyboard shortcut.** In VS Code, press `Ctrl+Shift+P`, type `Open Keyboard Shortcuts (JSON)`, and press Enter. Add this entry to the JSON array:
   ```json
   {
     "key": "ctrl+alt+c",
     "command": "workbench.action.tasks.runTask",
     "args": "Copy MD as HTML",
     "when": "editorLangId == markdown"
   }
   ```

4. **Reload VS Code.** Press `Ctrl+Shift+P`, type `Developer: Reload Window`, and press Enter. If the shortcut does not respond after reloading, close and reopen VS Code completely.

**How to use it:** open any `.md` file, press `Ctrl+Alt+C`, then paste directly into an Outlook email. The formatted text will paste correctly.



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



## Troubleshooting

**"pandoc is not installed or not in PATH"**
Pandoc is not installed, or PowerShell was not restarted after installation. Install Pandoc (see Step 1 above) and open a new PowerShell window.

**"Missing Python runtime: ...\\.venv\\Scripts\\python.exe" (when running `ld2d`)**
`ld2d` remaps legacy styles in-place to preserve cross-references, and needs the local virtual environment.
Create it from this repo root with:
```powershell
uv sync
```

**"d2m is not recognized as a command"**
The commands are not loaded. Either run `install_commands.ps1` again (see Step 2 above), or close and reopen PowerShell.

**"Input file not found"**
Check that the file name is correct and that you are in the right folder. Use `dir` to list the files in the current folder. If the file name contains spaces, make sure to wrap it in quotes: `d2m "My Contract.docx"`.

**`Ctrl+Alt+C` does nothing in a Markdown file**
VS Code may not have loaded the new task yet. Try a full restart (close and reopen VS Code, not just a reload window). You can also verify the task is available: press `Ctrl+Shift+P`, type `Tasks: Run Task`, and check that "Copy MD as HTML" appears in the list labelled "User".

**The VS Code task fails with a PowerShell error**
Open `%APPDATA%\Code\User\tasks.json` and confirm the task has `"type": "process"`, not `"type": "shell"`. The shell type routes the command through PowerShell, which parses the file path arguments and can fail when a path contains spaces.





# Technical Reference

Everything below is detailed technical documentation for developers and advanced users.

## Toolchain overview

| Script                               | Purpose                                         | Key dependencies                                                    |
| ------------------------------------ | ----------------------------------------------- | ------------------------------------------------------------------- |
| `commands/d2m.ps1`                   | DOCX to Markdown                                | `filters/docx_to_compact.lua`                                       |
| `commands/m2d.ps1`                   | Markdown to DOCX                                | `filters/compact_to_docx.lua`, `styles/flexup_template.docx`        |
| `commands/m2h.ps1`                   | Markdown to HTML (single self-contained file)   | `filters/compact_to_html.lua`, `styles/flexup_styles.css`           |
| `commands/ld2d.ps1`                  | Legacy DOCX to remapped DOCX                    | `commands/remap_legacy_contracts.py`, `styles/flexup_template.docx` |
| `commands/m2d-all.ps1`               | Batch Markdown to DOCX (all files in a folder)  | `m2d`                                                               |
| `commands/m2h-all.ps1`               | Batch Markdown to HTML (all files in a folder)  | `m2h`                                                               |
| `commands/d2m-all.ps1`               | Batch DOCX to Markdown (all files in a folder)  | `d2m`                                                               |
| `commands/commands.ps1`              | Loads all commands as PowerShell functions      | --                                                                  |
| `commands/install_commands.ps1`      | Adds commands loader to your PowerShell profile | --                                                                  |
| `commands/remap_legacy_contracts.py` | In-place legacy style remapping                 | `python-docx`                                                       |
| `commands/md2clip.py`                | Markdown → HTML clipboard (VS Code shortcut)    | Pandoc, PowerShell                                                  |

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

**Markdown to HTML:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\m2h.ps1 "SaaS contract.md"
powershell -ExecutionPolicy Bypass -File .\commands\m2h.ps1 "SaaS contract.md" "output.html"
```

**Legacy DOCX to remapped DOCX:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\ld2d.ps1 "legacy.docx"
powershell -ExecutionPolicy Bypass -File .\commands\ld2d.ps1 "legacy.docx" "output"
```

## Batch commands (direct invocation)

**Convert all Markdown files in a folder to DOCX:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\m2d-all.ps1
powershell -ExecutionPolicy Bypass -File .\commands\m2d-all.ps1 "C:\Contracts\Templates"
```

**Convert all DOCX files in a folder to Markdown:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\d2m-all.ps1
powershell -ExecutionPolicy Bypass -File .\commands\d2m-all.ps1 "C:\Contracts\Templates"
```

**Convert all Markdown files in a folder to HTML:**
```powershell
powershell -ExecutionPolicy Bypass -File .\commands\m2h-all.ps1
powershell -ExecutionPolicy Bypass -File .\commands\m2h-all.ps1 "C:\Contracts\Templates"
```

All batch commands process only the top-level folder (no recursion into subfolders). If an output file already exists, a counter is appended to the filename (e.g. `contract_1.docx`) to avoid overwriting.

## Output path behaviour

For single-file commands (`d2m`, `m2d`, `m2h`, `ld2d`, `dmd`, `mdm`):

- with no second argument, the output is created next to the input file
- with a second argument, relative output paths are resolved from the current PowerShell folder
- `.` writes to the current folder using the input file's base name
- `.\plop` writes to the current folder using `plop` as the base name
- an existing folder path, `.` / `..`, or a path ending in a slash is treated as an output folder
- absolute output paths are used as provided

Batch commands (`d2m-all`, `m2d-all`, `m2h-all`) still create each output file alongside its source file.

## Extension inference

All commands automatically add file extensions when omitted:

| Command   | Input default | Output default   |
| --------- | ------------- | ---------------- |
| `d2m`     | `.docx`       | `.md`            |
| `m2d`     | `.md`         | `.docx`          |
| `m2h`     | `.md`         | `.html`          |
| `ld2d`    | `.docx`       | `_remapped.docx` |
| `d2m-all` | all `.docx`   | `.md`            |
| `m2d-all` | all `.md`     | `.docx`          |
| `m2h-all` | all `.md`     | `.html`          |

Examples:
- `d2m "contract"` reads `contract.docx`, writes `contract.md`
- `m2d "contract" "final"` reads `contract.md`, writes `final.docx`
- `d2m "contract.docx" "."` reads `contract.docx`, writes `.\contract.md`
- `d2m "contract.docx" ".\plop"` reads `contract.docx`, writes `.\plop.md`
- `ld2d "legacy"` reads `legacy.docx`, writes `legacy_remapped.docx`
- `ld2d "legacy.docx" "clean"` reads `legacy.docx`, writes `clean.docx`

## Legacy style mapping

`ld2d` remaps styles directly in the DOCX file (in-place style remap), preserving Word cross-references and bookmarks.
It applies the following style conversions:

| Legacy style | New style |
| ------------ | --------- |
| Title        | Heading 1 |
| Heading 1    | Article 1 |
| Heading 2    | Article 2 |
| Heading 3    | Article 3 |
| Heading 4    | Article 4 |
| Heading 6    | Heading 4 |
| PublishedOn  | Published |
| ToDo         | Comments  |

## Python legacy remap

`ld2d` uses the Python script under the hood. You can also run it directly:

```powershell
# With activated virtual environment
python .\commands\remap_legacy_contracts.py legacy
python .\commands\remap_legacy_contracts.py legacy legacy_clean

# Without activating the virtual environment
uv run python .\commands\remap_legacy_contracts.py legacy
```

This requires Python 3.11+ and the `python-docx` package (managed via `pyproject.toml` and `uv`).

## Reference document

The file `styles/flexup_template.docx` is the Word reference template used by `m2d` and `ld2d` to produce correctly styled output. Keep this file aligned with your target Word style definitions. See the [style guide](styles/style_guide.md) and [style matrix](styles/style_matrix.md) for full style specifications.

`m2h` uses `styles/flexup_styles.css` instead -- the same Section/Article/Appendix numbering scheme, reimplemented as CSS counters targeting the classes that `filters/compact_to_html.lua` assigns.

## Related documentation

- [styles/style_guide.md](styles/style_guide.md) -- full specification of every paragraph style, numbering scheme, and Markdown mapping
- [styles/style_matrix.md](styles/style_matrix.md) -- compact reference table of all styles
- [styles/flexup_template.md](styles/flexup_template.md) -- example Markdown contract showing all styles in use
