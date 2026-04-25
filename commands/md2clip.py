#!/usr/bin/env python3
"""Convert a Markdown file to HTML and copy to Windows clipboard in CF_HTML format.

CF_HTML is the clipboard format that Outlook uses to distinguish real HTML
from plain text. Without it, Outlook may mangle paragraphs into bullet points.
"""

import sys
import os
import subprocess
import tempfile


def convert_md_to_html(filepath: str) -> str:
    result = subprocess.run(
        ["pandoc", "-f", "markdown", "-t", "html", filepath],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if result.returncode != 0:
        raise RuntimeError(f"Pandoc error: {result.stderr.strip()}")
    return result.stdout


def build_cf_html(html_fragment: str) -> str:
    """Wrap an HTML fragment in the Windows CF_HTML clipboard format.

    CF_HTML requires a plain-text header with byte offsets that point into the
    payload. The header itself counts toward those offsets, so we calculate
    everything after fixing the header length via zero-padded 10-digit numbers.
    """
    fragment = html_fragment.replace("\r\n", "\n").replace("\r", "\n")

    pre  = "<html>\r\n<body>\r\n<!--StartFragment-->"
    post = "<!--EndFragment-->\r\n</body>\r\n</html>"

    PAD = "0000000000"
    header = (
        f"Version:0.9\r\n"
        f"StartHTML:{PAD}\r\n"
        f"EndHTML:{PAD}\r\n"
        f"StartFragment:{PAD}\r\n"
        f"EndFragment:{PAD}\r\n"
    )

    h = len(header.encode("utf-8"))
    p = len(pre.encode("utf-8"))
    f = len(fragment.encode("utf-8"))
    q = len(post.encode("utf-8"))

    start_html     = h
    start_fragment = h + p
    end_fragment   = h + p + f
    end_html       = h + p + f + q

    header = header.replace(f"StartHTML:{PAD}",     f"StartHTML:{start_html:010d}",     1)
    header = header.replace(f"EndHTML:{PAD}",       f"EndHTML:{end_html:010d}",         1)
    header = header.replace(f"StartFragment:{PAD}", f"StartFragment:{start_fragment:010d}", 1)
    header = header.replace(f"EndFragment:{PAD}",   f"EndFragment:{end_fragment:010d}", 1)

    return header + pre + fragment + post


def copy_html_to_clipboard(cf_html: str) -> None:
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".html", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(cf_html)
        tmp_path = tmp.name.replace("\\", "/")

    try:
        ps = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            f"$h = [System.IO.File]::ReadAllText('{tmp_path}'); "
            "[System.Windows.Forms.Clipboard]::SetText"
            "($h, [System.Windows.Forms.TextDataFormat]::Html)"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(f"Clipboard write failed: {result.stderr.strip()}")
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: md2clip.py <file.md>", file=sys.stderr)
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.isfile(filepath):
        print(f"File not found: {filepath}", file=sys.stderr)
        sys.exit(1)

    html = convert_md_to_html(filepath)
    cf_html = build_cf_html(html)
    copy_html_to_clipboard(cf_html)
    print(f"Copied to clipboard: {os.path.basename(filepath)}")


if __name__ == "__main__":
    main()
