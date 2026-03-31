import sys, os, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from docx import Document
from docx.oxml.ns import qn

DOCX = r'c:\Files\Codes\Site\Alinggo_GreenHealing\Momo 網站使用（綠色療癒）\00-Week 1-18 ＧＨ EN_ZH版.docx'
DATA = r'c:\Files\Codes\Site\Alinggo_GreenHealing\data'

doc = Document(DOCX)
tmap = {t._element: t for t in doc.tables}
cw = None
wt = {}

for el in doc.element.body:
    if el.tag == qn('w:p'):
        tx = ''.join(n.text or '' for n in el.iter(qn('w:t'))).strip()
        norm = ''
        for c in tx:
            cp = ord(c)
            norm += chr(cp - 0xFEE0) if 0xFF01 <= cp <= 0xFF5E else c
        norm = norm.replace('\u3000', ' ')
        for ln in norm.split('\n'):
            m = re.match(r'^(?:\*{1,2})?Week\s+(\d{1,2})\b', ln.strip(), re.IGNORECASE)
            if m:
                v = int(m.group(1))
                if 1 <= v <= 18:
                    cw = v
    elif el.tag == qn('w:tbl') and el in tmap:
        if cw is None:
            cw = 1
        rows = [[c.text.strip() for c in r.cells] for r in tmap[el].rows]
        wt.setdefault(cw, []).append(rows)
        print("Table -> W{}: {} items, hdr={}".format(cw, len(rows)-1, rows[0] if rows else []))

for wn in sorted(wt):
    mp = os.path.join(DATA, "W{}".format(wn), 'content.md')
    if not os.path.exists(mp):
        continue
    with open(mp, 'r', encoding='utf-8') as f:
        md = f.read()
    parts = []
    for rows in wt[wn]:
        if not rows: continue
        h = rows[0]; n = len(h)
        parts.append('| ' + ' | '.join(h) + ' |')
        parts.append('| ' + ' | '.join(['---']*n) + ' |')
        for r in rows[1:]:
            while len(r) < n: r.append('')
            parts.append('| ' + ' | '.join(r) + ' |')
        parts.append('')
    tmd = '\n'.join(parts)
    lines = md.split('\n')
    ei = None
    for j, l in enumerate(lines):
        if re.match(r'^##\s*ESP\s*(Vocab|vocab)', l, re.IGNORECASE):
            ei = j; break
    if ei is not None:
        end = len(lines)
        for j in range(ei+1, len(lines)):
            if lines[j].startswith('## '): end = j; break
        md = '\n'.join(lines[:ei+1] + ['', tmd] + lines[end:])
    else:
        md = md.rstrip() + '\n\n## ESP Vocabulary\n\n' + tmd
    with open(mp, 'w', encoding='utf-8') as f:
        f.write(md)
    print("W{}: {} vocab items written".format(wn, sum(len(r)-1 for r in wt[wn])))

print("Done!")
