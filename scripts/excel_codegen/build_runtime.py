"""浏览器运行时构建:每流派产出一个自包含 ESM 模块。

dist/flows/{流派}.mjs 导出:
  meta            流派名/源文件/版本/生成时间/默认输出快照
  set(name, v)    按 flow-inputs.json 的语义字段名写输入
  setAll(obj) / reset() / fields()
  evaluate()      跑一遍整簿,返回 {总伤, ADPS, RDPS, 毕业率, 真气比例?}
  cell(ref)       调试:按 "表!坐标" 读任意格

语义规则:
  - set 只写期望侧槽位。RD 的联动格是 =期望!X 公式(自动跟随);
    带 rd_default 的字段(团队增益开关)RD 侧是 RDPS 的定义,不跟随。
  - reset 恢复所有字段槽(含 RD 侧定义值)。
  - 输出格按 期望 表 H 列标签定位,「DPS」归一为 ADPS(牵丝霖);
    总伤/ADPS/RDPS/毕业率 四项缺一即构建失败,真气比例可选。

用法: python build_runtime.py [upstream目录] [--out dist/flows]
"""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl.utils import column_index_from_string

from emit_js import RUNTIME
from optimize_js import OptimizingEmitter, default_inputs
from scan_inputs import flow_name
from workbook_graph import load_graph

ROOT = Path(__file__).parents[2]
FLOW_INPUTS = ROOT / "data" / "flow-inputs.json"

OUTPUT_LABELS = {
    "总伤": "总伤", "ADPS": "ADPS", "DPS": "ADPS",
    "RDPS": "RDPS", "毕业率": "毕业率", "真气比例": "真气比例",
}
REQUIRED_OUTPUTS = {"总伤", "ADPS", "RDPS", "毕业率"}


def scan_outputs(g) -> dict[str, tuple[str, str]]:
    """期望 H 列标签 -> (标准名, I 列公式格坐标)。"""
    out = {}
    for r in range(1, 27):
        label = g.values.get(("期望", f"H{r}"))
        if not isinstance(label, str):
            continue
        std = OUTPUT_LABELS.get(label.strip())
        if std and ("期望", f"I{r}") in g.nodes:
            out[std] = ("期望", f"I{r}")
    missing = REQUIRED_OUTPUTS - out.keys()
    if missing:
        raise RuntimeError(f"输出格缺失 {sorted(missing)}(H 列标签变了?)")
    return out


def build_flow(path: Path, fields_spec: list[dict], out_dir: Path) -> dict:
    g = load_graph(path)
    outputs = scan_outputs(g)

    inputs = set(default_inputs(g))
    for f in fields_spec:
        for sheet, coord in f["cells"].items():
            inputs.add((sheet, coord))

    em = OptimizingEmitter(g, inputs)

    # 字段 -> 槽位与默认值
    fields_js = {}
    for f in fields_spec:
        slots_all, writable = [], []
        for sheet, coord in f["cells"].items():
            slot = em.slot[(sheet, coord)]
            default = g.values.get((sheet, coord))
            if sheet == "RD" and "rd_default" in f:
                default = f["rd_default"]
            slots_all.append([slot, default])
            # RD 侧镜像若是独立字面量且无 rd_default(牵丝霖/破竹风的
            # 战斗时间等),语义上应跟随输入;带 rd_default 的是 RDPS
            # 定义(团队增益全关),不跟随
            if sheet == "期望" or "rd_default" not in f:
                writable.append(slot)
        fields_js[f["name"]] = {
            "kind": f["kind"],
            "default": f.get("default"),
            "set": writable,
            "all": slots_all,
            **({"note": f["note"]} if "note" in f else {}),
        }

    stmts = em.topo_statements()  # 先编译:填充 range_defs
    output_slots = {std: em.slot[key] for std, key in outputs.items()}
    output_defaults = {std: g.nodes[key].cached for std, key in outputs.items()}

    name = flow_name(path)
    version_m = re.search(r"(\d+(?:\.\d+)*)\.xlsx$", path.name)
    meta = {
        "flow": name,
        "sourceFile": path.name,
        "calcVersion": version_m.group(1) if version_m else None,
        "generatedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "defaults": output_defaults,
    }
    cell_slots = {f"{s}!{c}": i for (s, c), i in em.slot.items()}

    lines = [
        "// 由 scripts/excel_codegen/build_runtime.py 生成,请勿手改",
        RUNTIME,
        f"const x=new Array({len(em.slot)}).fill(null);",
        *em.literal_lines(),
        *em.range_defs,
        *em.chunk_lines(stmts),
        f"const FIELDS={json.dumps(fields_js, ensure_ascii=False)};",
        f"const OUTPUTS={json.dumps(output_slots, ensure_ascii=False)};",
        f"const CELLS={json.dumps(cell_slots, ensure_ascii=False)};",
        f"export const meta={json.dumps(meta, ensure_ascii=False)};",
        "export function fields(){return FIELDS;}",
        "export function set(name,value){const f=FIELDS[name];"
        'if(!f)throw new Error("未知字段: "+name);'
        "for(const s of f.set)x[s]=value;}",
        "export function setAll(obj){for(const k in obj)set(k,obj[k]);}",
        "export function reset(){for(const k in FIELDS)"
        "for(const [s,d] of FIELDS[k].all)x[s]=d;}",
        "export function evaluate(){EVAL();const out={};"
        "for(const k in OUTPUTS)out[k]=x[OUTPUTS[k]];return out;}",
        "export function cell(ref){const s=CELLS[ref];"
        'if(s===undefined)throw new Error("未知单元格: "+ref);return x[s];}',
    ]
    out_path = out_dir / f"{name}.mjs"
    out_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"{name}: {out_path.name} ({out_path.stat().st_size / 1024:.0f} KB) | {em.report()}")
    return {"flow": name, "module": out_path.name, **{k: meta[k] for k in ("sourceFile", "calcVersion", "defaults")}}


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    upstream = Path(args[0]) if args else ROOT / "upstream"
    out_dir = ROOT / "dist" / "flows"
    if "--out" in sys.argv:
        out_dir = Path(sys.argv[sys.argv.index("--out") + 1])
    out_dir.mkdir(parents=True, exist_ok=True)

    spec = json.loads(FLOW_INPUTS.read_text(encoding="utf-8"))["flows"]
    index = []
    for f in sorted(upstream.glob("*.xlsx")):
        name = flow_name(f)
        if name is None or name not in spec:
            continue
        index.append(build_flow(f, spec[name]["fields"], out_dir))
    (out_dir / "index.json").write_text(
        json.dumps(
            {"generatedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
             "flows": index},
            ensure_ascii=False, indent=2,
        ) + "\n",
        encoding="utf-8",
    )
    print(f"\n{len(index)} 个流派模块 -> {out_dir}")


if __name__ == "__main__":
    main()
