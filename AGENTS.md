# AGENTS.md — Setup guide for AI assistants

This file tells an AI assistant how to help a user set up the FlexUp Legal Styles toolbox on a **Windows 11** computer. Follow the steps below to get everything working from scratch.

---

## Prerequisites

Before starting, verify the required tools are installed. Run each check command in PowerShell:

| Tool          | Check command      | Required for                                |
| ------------- | ------------------ | ------------------------------------------- |
| Pandoc 3.x    | `pandoc --version` | All document conversions                    |
| Python 3.11+  | `python --version` | `ld2d` legacy remap, VS Code shortcut       |
| VS Code       | `code --version`   | Markdown editing and the clipboard shortcut |
| uv (optional) | `uv --version`     | Python virtual environment for `ld2d`       |

If a tool is missing:
- **Pandoc:** direct the user to `https://pandoc.org/installing.html` (Windows `.msi` installer). After installing, a new PowerShell window is needed.
- **Python:** install via `winget install Python.Python.3.11` or from `https://python.org/downloads`. Tick "Add Python to PATH" during installation.
- **uv:** `pip install uv` or `winget install astral-sh.uv`.
- **VS Code:** `winget install Microsoft.VisualStudioCode` or from `https://code.visualstudio.com`.

---

## Part 1 — PowerShell commands (`d2m`, `m2d`, `ld2d`, etc.)

### What the installer does

Running `commands/install_commands.ps1` performs four actions:

1. Creates a directory junction `C:\FlexUp` → the FlexUp shared documents root (the path is inferred from the script's own location — no hardcoded username)
2. Persists a `FLEXUP` environment variable at the Windows user level pointing to `C:\FlexUp`
3. Adds a clearly-delimited block to `$PROFILE` (typically `~\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1`) that dot-sources `commands/commands.ps1` in every new PowerShell session
4. Dot-sources the commands immediately so they are available in the current session

### How to run it

Navigate to the repo root (the `Styles` folder) in PowerShell and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\commands\install_commands.ps1
```

The installer is **idempotent** — re-running it is safe and replaces the existing profile block without duplicating it.

### What success looks like

```
Junction created for C:\FlexUp -> <path>
Profile configured: C:\Users\<user>\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
User environment variable configured: FLEXUP=C:\FlexUp
Close and reopen PowerShell so every new session picks up the updated profile.
```

### Verify after reopening PowerShell

```powershell
Get-Command d2m   # should resolve to a function
Get-Command m2d   # should resolve to a function
d2m "example\SaaS-GC.md"   # should produce example\SaaS-GC.docx
```

### Edge cases

- If `C:\FlexUp` already exists and points to the correct target, the installer skips junction creation.
- If `C:\FlexUp` already exists and points somewhere else, the installer throws. The user must delete `C:\FlexUp` first (this usually requires an elevated PowerShell: `cmd /c rmdir C:\FlexUp`).
- Profile updates require no elevation — only the junction step may need Administrator access.

### Optional: Python virtual environment (needed for `ld2d`)

`ld2d` remaps legacy Word styles in-place and requires `python-docx`. Set up the virtual environment once from the repo root:

```powershell
uv sync
```

This creates `.venv\` at the repo root. The `ld2d` command activates it automatically.

---

## Part 2 — VS Code Markdown-to-clipboard shortcut (`Ctrl+Alt+C`)

### What this does

When the user presses `Ctrl+Alt+C` while a `.md` file is active in VS Code:

1. VS Code runs `commands/md2clip.py` via Python, passing the current file path
2. Pandoc converts the Markdown source to a clean HTML fragment
3. The script wraps it in the Windows **CF_HTML** clipboard format
4. The HTML lands on the clipboard — ready to paste into Outlook with full formatting

CF_HTML is the specific Windows clipboard format that Office apps look for. Without it, Outlook treats pasted HTML as plain text and often mangles paragraphs into bullet points.

### Files to create

Three files need to be in place. The VS Code user data directory is `%APPDATA%\Code\User\` (typically `C:\Users\<name>\AppData\Roaming\Code\User\`).

---

#### File 1 — Python script

**Destination:** `%APPDATA%\Code\User\scripts\md2clip.py`

Create the `scripts` folder if it does not exist, then copy the file from the repo:

```powershell
$dest = "$env:APPDATA\Code\User\scripts"
if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
Copy-Item ".\commands\md2clip.py" "$dest\md2clip.py"
```

**Verify:** `python "$env:APPDATA\Code\User\scripts\md2clip.py"` should print `Usage: md2clip.py <file.md>`.

---

#### File 2 — VS Code user task

**Destination:** `%APPDATA%\Code\User\tasks.json`

If the file **does not exist**, create it with this exact content:

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

If the file **already exists**, read it and add the task object above into the existing `"tasks"` array. Do not replace the whole file.

> **Important:** use `"type": "process"`, not `"type": "shell"`. The shell type routes the command through PowerShell, which parses the arguments and fails on file paths that contain spaces.

---

#### File 3 — Keybinding

**Destination:** `%APPDATA%\Code\User\keybindings.json`

In VS Code, the user can open this file via `Ctrl+Shift+P` → "Open Keyboard Shortcuts (JSON)".

Add this entry to the JSON array:

```json
{
  "key": "ctrl+alt+c",
  "command": "workbench.action.tasks.runTask",
  "args": "Copy MD as HTML",
  "when": "editorLangId == markdown"
}
```

If the file does not exist yet, create it as:

```json
[
  {
    "key": "ctrl+alt+c",
    "command": "workbench.action.tasks.runTask",
    "args": "Copy MD as HTML",
    "when": "editorLangId == markdown"
  }
]
```

> **Note on JSONC:** VS Code uses JSONC (JSON with comments) for `keybindings.json`. The file may start with a `//` comment line — this is valid for VS Code but will cause Python's `json.load()` to fail. If you need to read the file programmatically, strip comment lines first or treat the file as plain text.

---

### After creating the files

Tell the user to reload VS Code: `Ctrl+Shift+P` → "Developer: Reload Window".

If the task still does not appear after a reload, a **full restart** (close and reopen VS Code) is needed the first time a user-level `tasks.json` is created.

---

### Verification

1. Open any `.md` file in VS Code
2. `Ctrl+Shift+P` → "Tasks: Run Task" → **"Copy MD as HTML"** should appear in the list, labelled "User"
3. Close the task picker, make sure the `.md` editor is focused, and press `Ctrl+Alt+C`
4. The terminal should briefly show: `Copied to clipboard: <filename>.md`
5. Paste into Outlook — formatting should be preserved (paragraphs as paragraphs, not bullets)

---

### Common issues

| Symptom                                                 | Cause                                  | Fix                                                                                                 |
| ------------------------------------------------------- | -------------------------------------- | --------------------------------------------------------------------------------------------------- |
| "Missing file specification after redirection operator" | `"type": "shell"` in tasks.json        | Change to `"type": "process"`                                                                       |
| `Ctrl+Alt+C` does nothing                               | VS Code hasn't loaded the new task yet | Fully restart VS Code                                                                               |
| "Copy MD as HTML" not in task list                      | tasks.json not yet picked up           | Fully restart VS Code                                                                               |
| Pandoc not found                                        | Pandoc not installed or not in PATH    | Install Pandoc; open a new VS Code window after install                                             |
| Paste into Outlook looks like plain text                | Script did not run successfully        | Check the terminal output for errors; run the script manually from PowerShell to see the full error |
