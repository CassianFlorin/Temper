"""输入格摘要:对各流派 xlsx 的 期望/RD 输入区做标签扫描。

产出 data/flow-inputs.json:语义字段名 -> 单元格坐标 的 per-flow 映射。
单元格坐标跨流派漂移(见 docs/xlsx-anatomy.md 第 6 节),但标签词汇
稳定,所以求解器/运行时按字段名注入,本文件负责翻译。

扫描规则(全部作用于 A1:I26 输入区,与 optimize_js.default_inputs
的边界一致):

  R1 面板矩阵   B1:E1 列头(最小攻击/最大攻击/穿透/伤害加成)×
                A2:A7 行头(外功/鸣金/裂石/牵丝/破竹/无相),
                字面量格 -> 面板.{行}.{列}
  R2 标量列     A8:A26 标签 + B 列字面量 -> {标签};紧随其后的
                无标签 B 格 -> {标签}#2(食物加成的 min/max 双行)
  R3 怪物/杂项  H 列标签 + I 列字面量 -> {标签}
  R4 团队增益   「团队增益」锚点下方 C 标签 + D √/× -> 团队增益.{标签}
  R5 开关对     E15:E26 标签 + F 列字面量 -> {标签}
  R6 纵向心法   任意「…心法」标签,取正下方文本 -> {标签}
  R7 属性/套装  F1:F3 标签 + G 列字面量 -> {标签}
  R8 武器键     F4:F7 中 G 侧为公式的文本键,排除固定词
                (单体奇术/群体奇术)-> 武器1/武器2
  R9 派生区表头 「实际属性」锚点的列头/行头,纯标签认领

R5/R6 对「标签在、值格是公式或空」的情况只认领标签不出字段(RD 表
的开关/心法多为 =期望!X 联动,属派生);值格为字面量才是输入。

收尾判定:区域内未认领字面量若不被任何公式引用(含区间展开),
归入 inert(上游遗留孤儿值,不阻塞);被引用却未认领的才进
unmapped —— unmapped 不清零不算扫描完成,新版式必须显式补规则。

RD 表同规则扫描,字段按名合并:默认值不同(团队增益在 RD 全关)
时记 rd_default。已知的上游半接线输入(鸣金虹 C21)以 KNOWN_EXTRAS
显式补充,来源见 data/constants.json 的 _resolved。
"""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl.utils import column_index_from_string, get_column_letter

from workbook_graph import WorkbookGraph, iter_refs, load_graph, resolve_ref

ROOT = Path(__file__).parents[2]
OUT = ROOT / "data" / "flow-inputs.json"

MAX_COL = 9   # A..I
MAX_ROW = 26

PANEL_COLS = {"最小攻击", "最大攻击", "穿透", "伤害加成"}
PANEL_ROWS = {"外功", "鸣金", "裂石", "牵丝", "破竹", "无相"}
FIXED_FG_KEYS = {"单体奇术", "群体奇术", "通用增伤"}

# 标签扫描无法发现的输入:上游半接线(空格但被公式消费)
KNOWN_EXTRAS = {
    "鸣金虹": [
        {
            "name": "第四心法",
            "kind": "text",
            "default": None,
            "cells": {"期望": "C21"},
            "block": "心法",
            "note": "上游 1.3 半接线:C21 为无标签空格,增益表以"
                    " IF(期望!$C$21=\"易水歌\") 消费;填「易水歌」激活整套 buff",
        }
    ],
}


def _coord(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"


class SheetScan:
    """单表扫描:literals[(col,row)] = 值,formulas = 公式格坐标集。"""

    def __init__(self, g: WorkbookGraph, sheet: str):
        self.sheet = sheet
        self.literals: dict[tuple[int, int], object] = {}
        self.formulas: set[tuple[int, int]] = set()
        for (s, coord), v in g.values.items():
            if s != sheet:
                continue
            m = re.fullmatch(r"([A-Z]+)(\d+)", coord)
            col, row = column_index_from_string(m.group(1)), int(m.group(2))
            if col <= MAX_COL and row <= MAX_ROW:
                self.literals[(col, row)] = v
        for (s, coord) in g.nodes:
            if s != sheet:
                continue
            m = re.fullmatch(r"([A-Z]+)(\d+)", coord)
            col, row = column_index_from_string(m.group(1)), int(m.group(2))
            if col <= MAX_COL and row <= MAX_ROW:
                self.formulas.add((col, row))
        self.claimed: set[tuple[int, int]] = set()
        self.fields: list[dict] = []

    # -- 工具

    def lit(self, col: int, row: int):
        return self.literals.get((col, row))

    def text(self, col: int, row: int) -> str | None:
        v = self.lit(col, row)
        return v if isinstance(v, str) else None

    def claim(self, *cells: tuple[int, int]) -> None:
        self.claimed.update(cells)

    def add_field(self, name: str, col: int, row: int, block: str) -> None:
        v = self.lit(col, row)
        kind = "toggle" if v in ("√", "×") else "text" if isinstance(v, str) else "number"
        self.fields.append(
            {"name": name, "kind": kind, "default": v,
             "cell": _coord(col, row), "block": block}
        )
        self.claim((col, row))

    # -- 规则

    def scan(self) -> None:
        self.r1_panel()
        self.r2_scalars()
        self.r3_monster()
        self.r4_team()
        self.r6_xinfa()   # 先于 R5:心法值格(断石之构等)是文本,别被 R5 当标签认领
        self.r5_ef_pairs()
        self.r7_fg_pairs()
        self.r8_weapon_keys()
        self.r9_derived_headers()

    def r1_panel(self) -> None:
        headers = {}
        for c in range(2, 6):  # B..E
            t = self.text(c, 1)
            if t in PANEL_COLS:
                headers[c] = t
                self.claim((c, 1))
        if not headers:
            return
        for r in range(2, 8):
            row_label = self.text(1, r)
            if row_label not in PANEL_ROWS:
                continue
            self.claim((1, r))
            for c, col_label in headers.items():
                if (c, r) in self.literals:
                    self.add_field(f"面板.{row_label}.{col_label}", c, r, "面板")

    def r2_scalars(self) -> None:
        last_label = None
        for r in range(8, MAX_ROW + 1):
            label = self.text(1, r)
            has_val = (2, r) in self.literals
            if label:
                self.claim((1, r))
                if has_val:
                    self.add_field(label, 2, r, "标量")
                    last_label = label
                else:
                    last_label = None  # 标签在但 B 是公式:派生量,不是输入
            elif has_val and last_label:
                self.add_field(f"{last_label}#2", 2, r, "标量")
                last_label = None
            else:
                last_label = None

    def r3_monster(self) -> None:
        for r in range(1, MAX_ROW + 1):
            label = self.text(8, r)
            if not label:
                continue
            self.claim((8, r))
            if (9, r) in self.literals:
                self.add_field(label, 9, r, "怪物/杂项")

    def r4_team(self) -> None:
        anchor = None
        for r in range(1, MAX_ROW + 1):
            if self.text(3, r) == "团队增益":
                anchor = r
                self.claim((3, r))
                break
        if anchor is None:
            return
        for r in range(anchor + 1, MAX_ROW + 1):
            label = self.text(3, r)
            v = self.lit(4, r)
            if label and v in ("√", "×"):
                self.claim((3, r))
                self.add_field(f"团队增益.{label}", 4, r, "团队增益")

    def r5_ef_pairs(self) -> None:
        for r in range(15, MAX_ROW + 1):
            label = self.text(5, r)
            if not label or (5, r) in self.claimed:
                continue
            self.claim((5, r))  # 值格为公式/空时是派生显示,标签仍认领
            if (6, r) in self.literals:
                self.add_field(label, 6, r, "开关")

    def r6_xinfa(self) -> None:
        for (c, r), v in sorted(self.literals.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            if not (isinstance(v, str) and v.endswith("心法")):
                continue
            self.claim((c, r))  # 下方为公式/空时(RD 联动)标签仍认领
            below = self.text(c, r + 1)
            if below and (c, r + 1) not in self.claimed:
                self.add_field(v, c, r + 1, "心法")

    def r7_fg_pairs(self) -> None:
        for r in range(1, 4):
            label = self.text(6, r)
            if label and (7, r) in self.literals:
                self.claim((6, r))
                self.add_field(label, 7, r, "属性/套装")
            elif label:
                self.claim((6, r))  # G 侧为公式:标签仍认领(如 通用增伤)

    def r8_weapon_keys(self) -> None:
        n = 0
        for r in range(4, 8):
            t = self.text(6, r)
            if t is None:
                continue
            if t in FIXED_FG_KEYS:
                self.claim((6, r))  # 固定词汇,查找键但非用户输入
                continue
            if (7, r) in self.formulas:
                n += 1
                self.add_field(f"武器{n}", 6, r, "武器")

    def r9_derived_headers(self) -> None:
        for r in range(1, MAX_ROW + 1):
            if self.text(3, r) != "实际属性":
                continue
            self.claim((3, r))
            for c in range(4, 8):  # D..G 列头
                if self.text(c, r + 1) in PANEL_COLS:
                    self.claim((c, r + 1))
            for rr in range(r + 2, r + 7):  # C 行头
                if self.text(3, rr) in PANEL_ROWS:
                    self.claim((3, rr))
            return

    def finish(self, referenced: set[tuple[int, int]]) -> tuple[list[dict], list[dict]]:
        """(unmapped 被引用未认领, inert 未被引用) 两桶。"""
        unmapped, inert = [], []
        for (c, r), v in sorted(self.literals.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            if (c, r) in self.claimed:
                continue
            item = {"cell": _coord(c, r), "value": v}
            (unmapped if (c, r) in referenced else inert).append(item)
        return unmapped, inert


def flow_name(path: Path) -> str | None:
    m = re.match(r"(.+?)110阶", path.stem)
    return m.group(1) if m else None


def region_references(g: WorkbookGraph) -> dict[str, set[tuple[int, int]]]:
    """期望/RD 扫描区内被任何公式(含区间展开)引用的格。"""
    covered: dict[str, set[tuple[int, int]]] = {"期望": set(), "RD": set()}
    seen_spans = set()
    for node in g.nodes.values():
        for ref in iter_refs(node.ast):
            span = resolve_ref(ref, node.sheet, g.sheet_extent)
            if span in seen_spans:
                continue
            seen_spans.add(span)
            sheet, r1, c1, r2, c2 = span
            if sheet not in covered:
                continue
            for r in range(r1, min(r2, MAX_ROW) + 1):
                for c in range(c1, min(c2, MAX_COL) + 1):
                    covered[sheet].add((c, r))
    return covered


def scan_workbook(path: Path) -> dict:
    g = load_graph(path)
    covered = region_references(g)
    scans = {sheet: SheetScan(g, sheet) for sheet in ("期望", "RD") if sheet in g.sheet_extent}
    for s in scans.values():
        s.scan()

    # 以 期望 为主合并 RD:同名字段记 RD 坐标与差异默认值
    main = scans["期望"]
    rd = scans.get("RD")
    rd_by_name = {f["name"]: f for f in rd.fields} if rd else {}
    fields = []
    for f in main.fields:
        entry = {
            "name": f["name"], "kind": f["kind"], "default": f["default"],
            "cells": {"期望": f["cell"]}, "block": f["block"],
        }
        r = rd_by_name.pop(f["name"], None)
        if r:
            entry["cells"]["RD"] = r["cell"]
            if r["default"] != f["default"]:
                entry["rd_default"] = r["default"]
        fields.append(entry)
    rd_only = [
        {"name": f["name"], "kind": f["kind"], "default": f["default"],
         "cells": {"RD": f["cell"]}, "block": f["block"]}
        for f in rd_by_name.values()
    ]

    name = flow_name(path)
    for extra in KNOWN_EXTRAS.get(name or "", []):
        fields.append(dict(extra))

    unmapped, inert = {}, {}
    for sheet, s in scans.items():
        um, inrt = s.finish(covered.get(sheet, set()))
        if um:
            unmapped[sheet] = um
        if inrt:
            inert[sheet] = inrt

    return {
        "file": path.name,
        "fields": fields + rd_only,
        "unmapped": unmapped,
        "inert": inert,
    }


def main() -> None:
    upstream = Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / "upstream")
    flows = {}
    total_unmapped = 0
    for f in sorted(upstream.glob("*.xlsx")):
        name = flow_name(f)
        if name is None:  # DIY 计算器:版式不同且非流派,不在本映射范围
            continue
        result = scan_workbook(f)
        flows[name] = result
        n_un = sum(len(v) for v in result["unmapped"].values())
        n_inert = sum(len(v) for v in result["inert"].values())
        total_unmapped += n_un
        print(f"{name}: 字段 {len(result['fields'])}, 未认领 {n_un}, 惰性 {n_inert}")
        for sheet, items in result["unmapped"].items():
            for it in items:
                print(f"    未认领 [{sheet}!{it['cell']}] {it['value']!r}")
        for sheet, items in result["inert"].items():
            for it in items:
                print(f"    惰性   [{sheet}!{it['cell']}] {it['value']!r}")

    OUT.write_text(
        json.dumps(
            {
                "_meta": {
                    "generatedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    "generator": "scripts/excel_codegen/scan_inputs.py",
                    "note": "字段名 -> 期望/RD 单元格坐标。unmapped 必须为空或"
                            "逐项有人工解释;新版式出现未认领格时先补规则再发布。",
                },
                "flows": flows,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    print(f"\n写入 {OUT.relative_to(ROOT)};未认领合计 {total_unmapped}")
    if total_unmapped:
        sys.exit(1)  # 新版式出现未认领输入格:先补规则再发布


if __name__ == "__main__":
    main()
