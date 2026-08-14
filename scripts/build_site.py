"""站点组装:dist/site = web 源码 + 流派运行时模块 + 数据文件。

前置:scripts/excel_codegen/build_runtime.py 已产出 dist/flows/。
用法: python scripts/build_site.py
"""

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).parents[1]
SITE = ROOT / "dist" / "site"
FLOWS = ROOT / "dist" / "flows"


def main() -> None:
    if not (FLOWS / "index.json").exists():
        sys.exit("dist/flows 不存在,先跑 scripts/excel_codegen/build_runtime.py")

    if SITE.exists():
        shutil.rmtree(SITE)
    SITE.mkdir(parents=True)

    for f in (ROOT / "web").iterdir():
        if f.is_file():
            shutil.copy2(f, SITE / f.name)
    shutil.copytree(FLOWS, SITE / "flows")
    (SITE / "data").mkdir()
    for name in ("constants.json", "upstream-manifest.json", "flow-inputs.json"):
        shutil.copy2(ROOT / "data" / name, SITE / "data" / name)

    total = sum(p.stat().st_size for p in SITE.rglob("*") if p.is_file())
    print(f"dist/site 组装完成,{total / 1024 / 1024:.1f} MB")


if __name__ == "__main__":
    main()
