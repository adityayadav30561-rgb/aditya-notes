#!/usr/bin/env python3
"""Print numbered questions (with section headers): number.py <source.md> [from] [to]."""
import sys

src = sys.argv[1]
lo = int(sys.argv[2]) if len(sys.argv) > 2 else 1
hi = int(sys.argv[3]) if len(sys.argv) > 3 else 10**9
n, sec, shown = 0, "", None
for line in open(src, encoding="utf-8"):
    if line.startswith("## "):
        sec = line[3:].strip()
    elif line.startswith("- "):
        n += 1
        if lo <= n <= hi:
            if sec != shown:
                print(f"## {sec}")
                shown = sec
            print(f"{n}|{line[2:].split(' :: ')[0].strip()}")
