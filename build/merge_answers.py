#!/usr/bin/env python3
"""Merge answers into a source file: merge_answers.py <source.md> <answers.txt>...

Each answers file has lines "N|answer", where N is the question's number in the
source (as printed by number.py). The answer is stored on the question line
after " :: ", replacing any existing answer.
"""
import sys

src, files = sys.argv[1], sys.argv[2:]
answers = {}
for f in files:
    for line in open(f, encoding="utf-8"):
        line = line.rstrip("\n")
        if "|" in line and line.split("|", 1)[0].strip().isdigit():
            num, ans = line.split("|", 1)
            answers[int(num)] = ans.strip()

out, n = [], 0
for line in open(src, encoding="utf-8"):
    if line.startswith("- "):
        n += 1
        if n in answers:
            q = line[2:].split(" :: ")[0].rstrip("\n").strip()
            line = f"- {q} :: {answers[n]}\n" if answers[n] else f"- {q}\n"
    out.append(line)
open(src, "w", encoding="utf-8").write("".join(out))
print(f"{src}: {len(answers)} answers merged, {n} questions")
