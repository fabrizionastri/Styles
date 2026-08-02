"""Rewrite two-column grid "definitions" tables as ::: Definitions blocks.

Each source line inside a cell was authored as its own Word paragraph, so each
line becomes its own block: a "- " run becomes one list, ::: fences keep their
div, everything else becomes a paragraph.
"""
import re
import sys


def split_row(line, seps):
    return [line[a + 1:b].rstrip() for a, b in zip(seps, seps[1:])]


def cell_blocks(lines):
    """Group a cell's raw lines into markdown blocks."""
    blocks, buf, mode = [], [], None

    def flush():
        nonlocal buf, mode
        if buf:
            blocks.append((mode, buf))
        buf, mode = [], None

    for line in lines:
        s = line.strip()
        if not s:
            flush()
            continue
        if s.startswith(':::'):
            flush()
            blocks.append(('fence', [s]))
            continue
        want = 'list' if s.startswith('- ') else 'para'
        if want != mode or want == 'para':
            flush()
        mode = want
        buf.append(s)
    flush()
    return blocks


def render_definition(blocks, indent='    '):
    out, in_div = [], False
    for kind, lines in blocks:
        if kind == 'fence':
            if in_div:
                out.append(indent + ':::')
                out.append('')
                in_div = False
            else:
                out.append(indent + lines[0])
                in_div = True
            continue
        body = [indent + l for l in lines]
        out.extend(body)
        if not in_div:
            out.append('')
    while out and out[-1] == '':
        out.pop()
    return out


def convert(text):
    lines = text.split('\n')
    out, i = [], 0
    while i < len(lines):
        line = lines[i]
        if not re.match(r'^\+[-=+]+\+$', line.replace(' ', '')) or line.count('+') != 3:
            out.append(line)
            i += 1
            continue

        seps = [m.start() for m in re.finditer(r'\+', line)]
        rows, j = [], i + 1
        current = [[], []]
        while j < len(lines) and (lines[j].startswith('|') or lines[j].startswith('+')):
            if lines[j].startswith('+'):
                rows.append(current)
                current = [[], []]
            else:
                a, b = split_row(lines[j], seps)
                current[0].append(a)
                current[1].append(b)
            j += 1
        if current[0] or current[1]:
            rows.append(current)
        rows = [r for r in rows if any(c.strip() for c in r[0] + r[1])]

        # Only a term/definition table qualifies: every first cell must be a
        # single non-empty line.
        terms_ok = all(
            len([x for x in r[0] if x.strip()]) == 1 for r in rows
        ) and all(any(x.strip() for x in r[1]) for r in rows)
        if not terms_ok or len(rows) < 2:
            out.extend(lines[i:j])
            i = j
            continue

        out.append('::: Definitions')
        out.append('')
        for r in rows:
            term = next(x.strip() for x in r[0] if x.strip())
            out.append(term)
            defn = render_definition(cell_blocks(r[1]))
            if defn:
                first = defn[0].lstrip()
                out.append(':   ' + first)
                out.extend(defn[1:])
            out.append('')
        out.append(':::')
        i = j
    return '\n'.join(out)


src = sys.argv[1]
dst = sys.argv[2]
text = open(src, encoding='utf-8').read()
open(dst, 'w', encoding='utf-8', newline='\n').write(convert(text))
print('wrote', dst)
