# -*- coding: utf-8 -*-
import os, re
from collections import OrderedDict

root = r"F:\CursorProject_Python\_bes_tr_analysis"
pat = re.compile(
    r"^(?:Public |Private |Friend )?(Sub|Function|Property Get|Property Let|Property Set)\s+([A-Za-z0-9_]+)",
    re.M,
)
out = []
total_subs = 0
for fn in sorted(os.listdir(root)):
    if not fn.endswith(".bas"):
        continue
    path = os.path.join(root, fn)
    with open(path, encoding="utf-8", errors="replace") as f:
        text = f.read()
    lines = text.splitlines()
    nlines = len(lines)
    procs = pat.findall(text)
    names = [p[1] for p in procs]
    total_subs += len(names)
    vbname = ""
    m = re.search(r'Attribute VB_Name = "([^"]+)"', text)
    if m:
        vbname = m.group(1)
    kind = "Sheet" if fn.endswith(".cls.bas") and vbname in {"T1","T2","T3","T4","T5","T6","T7","T8","T9","W1","W4","DieseArbeitsmappe","T1"} else (
        "Form" if fn.endswith(".frm.bas") else ("Sheet" if fn.endswith(".cls.bas") else "Module")
    )
    out.append((vbname or fn, kind, nlines, len(names), names[:25], fn))

with open(os.path.join(root, "_vba_inventory.txt"), "w", encoding="utf-8") as f:
    f.write(f"TOTAL files {len(out)} procs {total_subs}\n")
    for name, kind, nlines, nprocs, names, fn in sorted(out, key=lambda x: (x[1], x[0].lower())):
        f.write(f"\n[{kind}] {name}  lines={nlines} procs={nprocs}  file={fn}\n")
        for n in names:
            f.write(f"  - {n}\n")
        if nprocs > 25:
            f.write(f"  ... +{nprocs-25} more\n")
print("ok", total_subs)
