"""运行时模块验收:接线级检查(单元格级等价已由 verify_equiv 保证)。

对 dist/flows/ 每个模块:
  1. reset + evaluate:输出与 xlsx 缓存值一致(相对容忍 1e-9)
  2. set(战斗时间, ×1.25):ADPS 与 RDPS 必须都变(RD 联动生效)
  3. reset + evaluate:与第 1 步位级一致(reset 完整还原)

用法: python verify_runtime.py [dist/flows]
"""

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parents[2]

RUNNER = """
const [modPath] = process.argv.slice(1);
const m = await import("file://" + modPath);
m.reset();
const r1 = m.evaluate();
const t = m.fields()["战斗时间"].default;
m.set("战斗时间", t * 1.25);
const r2 = m.evaluate();
m.reset();
const r3 = m.evaluate();
console.log(JSON.stringify({r1, r2, r3, meta: m.meta}));
"""


def close(a: float, b: float) -> bool:
    return abs(a - b) <= 1e-9 * max(1.0, abs(b))


def verify(mod: Path) -> bool:
    proc = subprocess.run(
        ["node", "--input-type=module", "-e", RUNNER, "--", str(mod)],
        capture_output=True, text=True, timeout=120,
    )
    if proc.returncode != 0:
        print(f"FAIL {mod.name}: node 失败\n{proc.stderr[:1200]}")
        return False
    r = json.loads(proc.stdout)
    r1, r2, r3, meta = r["r1"], r["r2"], r["r3"], r["meta"]
    problems = []
    for k, want in meta["defaults"].items():
        if not close(r1[k], want):
            problems.append(f"默认输出 {k}: 得 {r1[k]!r} 期望 {want!r}")
    for k in ("ADPS", "RDPS"):
        if close(r2[k], r1[k]):
            problems.append(f"扰动战斗时间后 {k} 未变化(RD 联动断了?)")
    for k in r1:
        if r3[k] != r1[k]:
            problems.append(f"reset 后 {k} 未还原: {r3[k]!r} != {r1[k]!r}")
    status = "OK  " if not problems else "FAIL"
    print(f"{status}{mod.name}: 毕业率默认 {r1['毕业率']:.6f}")
    for p in problems:
        print(f"    {p}")
    return not problems


def main() -> None:
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "dist" / "flows"
    mods = sorted(out_dir.glob("*.mjs"))
    if not mods:
        sys.exit(f"{out_dir} 下没有模块,先跑 build_runtime.py")
    if not all([verify(m) for m in mods]):
        sys.exit(1)


if __name__ == "__main__":
    main()
