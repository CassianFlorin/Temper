"""AND 回归:普通/优化发射器均须保留 Excel 逻辑与错误传播语义。"""

import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from emit_js import Emitter
from formula_parser import parse_formula
from optimize_js import OptimizingEmitter
from workbook_graph import CellNode, WorkbookGraph


class LogicalTests(unittest.TestCase):
    def test_upstream_weapon_and_inner_skill_switches(self):
        # 裂石钧 3.1 原始公式:切换武器/心法后,对应加成必须独立开关。
        formulas = {
            "D2": '=58.68+IF(AND($E$24="唐刀",$E$22="征人归"),5.1,0)',
            "B11": '=9.2%+IF(AND($E$24="唐刀",$E$22="断石之构"),4.1%,0)',
            "I8": '=IF($E$24="陌刀",89.4,88.1)',
        }
        nodes = {
            ("期望", coord): CellNode(
                sheet="期望", coord=coord, formula=formula,
                is_array=False, ast=parse_formula(formula), cached=None,
            )
            for coord, formula in formulas.items()
        }
        values = {("期望", "E24"): "唐刀", ("期望", "E22"): "征人归"}
        graph = WorkbookGraph(
            path=Path("logical.xlsx"), nodes=nodes, values=values,
            sheet_extent={"期望": (24, 9)}, topo=list(nodes),
        )
        cases = [
            ("唐刀", "征人归", 63.78, 0.092),
            ("唐刀", "断石之构", 58.68, 0.133),
            ("陌刀", "征人归", 58.68, 0.092),
            ("陌刀", "断石之构", 58.68, 0.092),
        ]
        with tempfile.TemporaryDirectory(prefix="temper_switches_") as temp:
            for emitter in (Emitter, OptimizingEmitter):
                path = Path(temp) / "switches.js"
                path.write_text(emitter(graph, inputs=set(values)).emit(), encoding="utf-8")
                for weapon, skill, attack, crit in cases:
                    with self.subTest(emitter=emitter.__name__, weapon=weapon, skill=skill):
                        overrides = {"期望!E24": weapon, "期望!E22": skill}
                        result = subprocess.run(
                            ["node", str(path), json.dumps(overrides, ensure_ascii=False)],
                            capture_output=True, text=True, check=True, timeout=30,
                        )
                        got = json.loads(result.stdout)
                        self.assertAlmostEqual(got["期望!D2"], attack)
                        self.assertAlmostEqual(got["期望!B11"], crit)
                        self.assertAlmostEqual(
                            got["期望!I8"], 89.4 if weapon == "陌刀" else 88.1,
                        )

    def test_and_in_both_emitters(self):
        # 比较表达式对应上游唐刀/心法开关;引用用例覆盖文本、空格与错误。
        cases = [
            ('=AND(TRUE,TRUE)', True),
            ('=AND(TRUE,FALSE)', False),
            ('=AND(2,-1)', True),
            ('=AND(2,0)', False),
            ('=IF(AND(A1>0,A2=0),"match","miss")', "match"),
            ('=_xlfn.AND(TRUE,A1>0)', True),
            ('=AND(A1:A4)', False),
            ('=AND(A3:A4)', True),
            ('=AND((A3),TRUE)', True),
            ('=AND(A6,TRUE)', True),
            ('=AND(A3:A3)', {"__err": "#VALUE!"}),
            ('=AND(A6:A6)', {"__err": "#VALUE!"}),
            ('=AND(TRUE,"invalid")', {"__err": "#VALUE!"}),
            ('=AND(FALSE,#N/A)', {"__err": "#N/A"}),
            ('=AND(FALSE,B1)', {"__err": "#N/A"}),
            ('=AND(FALSE,B1:B2)', {"__err": "#N/A"}),
            ('=IFERROR(AND(FALSE,#N/A),42)', 42),
            ('=IFERROR(AND(A3:A3),42)', 42),
            ('=IFERROR(AND(FALSE,B1),42)', 42),
        ]
        formulas = [("B1", "=#N/A"), ("B2", "=#REF!")]
        formulas += [(f"C{i}", f) for i, (f, _) in enumerate(cases, 1)]
        nodes = {
            ("期望", coord): CellNode(
                sheet="期望", coord=coord, formula=formula,
                is_array=False, ast=parse_formula(formula), cached=None,
            )
            for coord, formula in formulas
        }
        graph = WorkbookGraph(
            path=Path("logical.xlsx"), nodes=nodes,
            values={
                ("期望", "A1"): 2, ("期望", "A2"): 0,
                ("期望", "A3"): "text", ("期望", "A4"): True,
            },
            sheet_extent={"期望": (len(cases), 3)}, topo=list(nodes),
        )
        expected = {f"期望!C{i}": v for i, (_, v) in enumerate(cases, 1)}
        node = shutil.which("node")
        self.assertIsNotNone(node, "Node.js is required for runtime tests")
        with tempfile.TemporaryDirectory(prefix="temper_logical_") as temp:
            for emitter in (Emitter, OptimizingEmitter):
                with self.subTest(emitter=emitter.__name__):
                    path = Path(temp) / "logical.js"
                    path.write_text(emitter(graph, inputs=set()).emit(), encoding="utf-8")
                    result = subprocess.run(
                        [node, str(path)], capture_output=True, text=True,
                        check=True, timeout=30,
                    )
                    got = json.loads(result.stdout)
                    self.assertEqual({k: got[k] for k in expected}, expected)


if __name__ == "__main__":
    unittest.main()
