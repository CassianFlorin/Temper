"""优化 pass 验收:金标准 + 输入扰动等价性模糊测试。

两道关:
  1. 默认输入下,优化版输出仍逐格通过缓存值金标准(同 verify_golden)。
  2. 随机扰动输入格(数值 ×U(0.5,1.5),√/× 以 0.3 概率翻转,心法在
     候选间不动 —— 文本输入保持原值域),优化版与未优化版逐格比对,
     要求位级相等:优化只做查表折叠,不改变任何算术结构与求值顺序。

用法: python verify_equiv.py [upstream目录|单个xlsx] [--trials N]
"""

import json
import random
import subprocess
import sys
import tempfile
from pathlib import Path

from emit_js import Emitter
from optimize_js import OptimizingEmitter, default_inputs
from verify_golden import compare
from workbook_graph import load_graph

TRIALS = 3


def run_node(js_path: Path, overrides: dict | None) -> dict:
    cmd = ["node", "--max-old-space-size=4096", str(js_path)]
    if overrides is not None:
        cmd.append(json.dumps(overrides, ensure_ascii=False))
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if proc.returncode != 0:
        raise RuntimeError(f"node 失败: {proc.stderr[:1500]}")
    return json.loads(proc.stdout)


def make_overrides(g, inputs, rng) -> dict:
    ovr = {}
    for sheet, coord in sorted(inputs):
        v = g.values[(sheet, coord)]
        k = f"{sheet}!{coord}"
        if isinstance(v, bool):
            continue
        if isinstance(v, (int, float)):
            ovr[k] = round(v * rng.uniform(0.5, 1.5), 6)
        elif v == "√" and rng.random() < 0.3:
            ovr[k] = "×"
        elif v == "×" and rng.random() < 0.3:
            ovr[k] = "√"
    return ovr


def verify(path: Path, trials: int) -> bool:
    g = load_graph(path)
    inputs = default_inputs(g)
    base_src = Emitter(g, inputs).emit()
    opt_em = OptimizingEmitter(g)
    opt_src = opt_em.emit()

    tmpdir = Path(tempfile.mkdtemp(prefix="temper_equiv_"))
    base_js = tmpdir / "base.js"
    opt_js = tmpdir / "opt.js"
    base_js.write_text(base_src, encoding="utf-8")
    opt_js.write_text(opt_src, encoding="utf-8")

    size_ratio = len(opt_src) / len(base_src)
    keys = [f"{s}!{c}" for s, c in g.nodes]

    # 关卡 1:优化版过金标准
    got = run_node(opt_js, None)
    golden_fail = 0
    for (sheet, coord), node in g.nodes.items():
        if compare(node.cached, got.get(f"{sheet}!{coord}")) == "FAIL":
            golden_fail += 1
            if golden_fail <= 5:
                print(f"  金标准 FAIL {sheet}!{coord}: 缓存 {node.cached!r} 求值 {got.get(f'{sheet}!{coord}')!r}")

    # 关卡 2:扰动等价
    rng = random.Random(20260814)
    fuzz_fail = 0
    for t in range(trials):
        ovr = make_overrides(g, inputs, rng)
        a = run_node(base_js, ovr)
        b = run_node(opt_js, ovr)
        for k in keys:
            if a[k] != b[k]:
                fuzz_fail += 1
                if fuzz_fail <= 5:
                    print(f"  扰动#{t} 不等价 {k}: 未优化 {a[k]!r} 优化 {b[k]!r}")

    status = "OK  " if not (golden_fail or fuzz_fail) else "FAIL"
    print(
        f"{status}{path.name}: 输入格 {len(inputs)}, 体积比 {size_ratio:.2f}, "
        f"金标准FAIL {golden_fail}, 扰动不等价 {fuzz_fail} ({trials} 轮) | {opt_em.report()}"
    )
    return not (golden_fail or fuzz_fail)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    trials = TRIALS
    if "--trials" in sys.argv:
        trials = int(sys.argv[sys.argv.index("--trials") + 1])
        args = [a for a in args if a != str(trials)]
    target = Path(args[0]) if args else Path(__file__).parents[2] / "upstream"
    files = [target] if target.is_file() else sorted(target.glob("*.xlsx"))

    ok = all([verify(f, trials) for f in files])
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
