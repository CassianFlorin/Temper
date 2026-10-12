"""上游版式回归:纵向模式必须接线,由公式派生的时间不能伪装成输入。"""

import unittest
from pathlib import Path

from formula_parser import parse_formula
from scan_inputs import SheetScan
from workbook_graph import CellNode, WorkbookGraph


class InputScanTests(unittest.TestCase):
    def test_vertical_mode_and_derived_time(self):
        formula = '=IF($E$24="陌刀",89.4,88.1)'
        graph = WorkbookGraph(
            path=Path("inputs.xlsx"),
            values={
                ("期望", "E21"): "唐刀第四心法",
                ("期望", "E22"): "断石之构",
                ("期望", "E23"): "输出模式",
                ("期望", "E24"): "陌刀",
                ("期望", "H8"): "战斗时间",
            },
            nodes={
                ("期望", "I8"): CellNode(
                    sheet="期望", coord="I8", formula=formula,
                    is_array=False, ast=parse_formula(formula), cached=89.4,
                )
            },
            sheet_extent={"期望": (24, 9)},
        )
        scan = SheetScan(graph, "期望")
        scan.scan()
        fields = {f["name"]: f for f in scan.fields}
        self.assertEqual(fields["输出模式"]["cell"], "E24")
        self.assertEqual(fields["输出模式"]["default"], "陌刀")
        self.assertEqual(fields["唐刀第四心法"]["cell"], "E22")
        self.assertNotIn("战斗时间", fields)
        # E24 被公式消费,必须作为输入接线,不能仅被 R5 认领成标签。
        unmapped, _ = scan.finish({(5, 24)})
        self.assertEqual(unmapped, [])


if __name__ == "__main__":
    unittest.main()
