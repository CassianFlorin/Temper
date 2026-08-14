"""解析器语料回归:把 upstream/ 下所有 xlsx 的全部公式解析后回写,与原文比对。

用法: python parse_corpus.py [upstream目录]

三级判定:
  exact  回写与原文逐字符一致
  tokeq  回写与原文 token 序列一致(仅空白等无语义差异)
  FAIL   解析失败或 token 序列不一致
"""

import sys
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

from formula_parser import parse_formula, tokenize, FormulaError

UPSTREAM = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parents[2] / "upstream")


def collect(path: Path):
    wb = openpyxl.load_workbook(path, data_only=False)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, ArrayFormula):
                    yield ws.title, c.coordinate, v.text
                elif isinstance(v, str) and v.startswith("="):
                    yield ws.title, c.coordinate, v


def main():
    files = sorted(UPSTREAM.glob("*.xlsx"))
    if not files:
        sys.exit(f"{UPSTREAM} 下没有 xlsx")

    grand = Counter()
    failures = []
    for f in files:
        seen = {}
        for sheet, coord, text in collect(f):
            seen.setdefault(text, (sheet, coord))
        stats = Counter()
        for text, (sheet, coord) in seen.items():
            try:
                out = "=" + parse_formula(text).to_formula()
                if out == text:
                    stats["exact"] += 1
                elif [t.text for t in tokenize(out[1:])] == [t.text for t in tokenize(text.lstrip("="))]:
                    stats["tokeq"] += 1
                else:
                    stats["FAIL"] += 1
                    failures.append((f.name, sheet, coord, text, out))
            except FormulaError as e:
                stats["FAIL"] += 1
                failures.append((f.name, sheet, coord, text, f"<{e}>"))
        grand.update(stats)
        total = sum(stats.values())
        print(f"{f.name}: 唯一公式 {total}, exact {stats['exact']}, tokeq {stats['tokeq']}, FAIL {stats['FAIL']}")

    total = sum(grand.values())
    print(f"\n合计: 唯一公式 {total}, exact {grand['exact']}, tokeq {grand['tokeq']}, FAIL {grand['FAIL']}")
    for fn, sheet, coord, text, out in failures[:10]:
        print(f"\nFAIL {fn} [{sheet}!{coord}]\n  原文: {text}\n  回写: {out}")
    if grand["FAIL"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
