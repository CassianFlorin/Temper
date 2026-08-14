"""金标准比对:生成 JS → node 执行 → 与 xlsx 缓存值逐格比对。

用法: python verify_golden.py [upstream目录|单个xlsx] [--keep]

判定:
  exact  数值位级相等 / 字符串·布尔全等
  close  数值相对误差 ≤ 1e-9(浮点运算顺序差异的容忍带)
  FAIL   其余(含 __err)

任何 FAIL 退出码 1。--keep 保留生成的 js 于 build/ 供排查。
"""

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from workbook_graph import load_graph
from emit_js import Emitter

REL_TOL = 1e-9


def compare(cached, got):
    if isinstance(got, dict) and "__err" in got:
        return "FAIL"
    if isinstance(cached, bool) or isinstance(got, bool):
        return "exact" if cached is got else "FAIL"
    if isinstance(cached, (int, float)) and isinstance(got, (int, float)):
        if float(cached) == float(got):
            return "exact"
        if abs(got - cached) <= REL_TOL * max(1.0, abs(cached)):
            return "close"
        return "FAIL"
    if isinstance(cached, str) and isinstance(got, str):
        return "exact" if cached == got else "FAIL"
    # 类型不一致(如公式结果为空 vs 缓存 0)
    if cached in (0, "") and got is None:
        return "close"
    return "FAIL"


def verify(path: Path, keep_dir: Path | None):
    g = load_graph(path)
    src = Emitter(g).emit()
    if keep_dir:
        keep_dir.mkdir(exist_ok=True)
        js_path = keep_dir / (path.stem + ".gen.js")
        js_path.write_text(src, encoding="utf-8")
    else:
        tmp = tempfile.NamedTemporaryFile(
            "w", suffix=".js", delete=False, encoding="utf-8"
        )
        tmp.write(src)
        tmp.close()
        js_path = Path(tmp.name)

    try:
        proc = subprocess.run(
            ["node", "--max-old-space-size=4096", str(js_path)],
            capture_output=True, text=True, timeout=300,
        )
        if proc.returncode != 0:
            print(f"FAIL {path.name}: node 执行失败\n{proc.stderr[:2000]}")
            return None
        got = json.loads(proc.stdout)
    finally:
        if not keep_dir:
            js_path.unlink(missing_ok=True)

    stats = {"exact": 0, "close": 0, "FAIL": 0}
    failures = []
    for (sheet, coord), node in g.nodes.items():
        k = f"{sheet}!{coord}"
        verdict = compare(node.cached, got.get(k))
        stats[verdict] += 1
        if verdict == "FAIL":
            failures.append((k, node.formula, node.cached, got.get(k)))
    total = sum(stats.values())
    print(
        f"{path.name}: 公式格 {total}, exact {stats['exact']}, "
        f"close {stats['close']}, FAIL {stats['FAIL']}"
    )
    for k, formula, cached, gv in failures[:8]:
        print(f"  FAIL {k}\n    公式: {formula[:160]}\n    缓存: {cached!r}\n    求值: {gv!r}")
    return stats


def main():
    args = [a for a in sys.argv[1:] if a != "--keep"]
    keep = "--keep" in sys.argv
    target = Path(args[0]) if args else Path(__file__).parents[2] / "upstream"
    files = [target] if target.is_file() else sorted(target.glob("*.xlsx"))
    keep_dir = Path(__file__).parents[2] / "build" if keep else None

    grand = {"exact": 0, "close": 0, "FAIL": 0}
    broken = False
    for f in files:
        stats = verify(f, keep_dir)
        if stats is None:
            broken = True
            continue
        for k in grand:
            grand[k] += stats[k]
    total = sum(grand.values())
    print(
        f"\n合计: 公式格 {total}, exact {grand['exact']}, "
        f"close {grand['close']}, FAIL {grand['FAIL']}"
    )
    if broken or grand["FAIL"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
