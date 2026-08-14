"""Excel 公式解析器:公式文本 → AST。

覆盖上游社区计算器实测用到的语法面(见 docs/xlsx-anatomy.md 第 4 节):
IF / VLOOKUP / SUM / _xlfn.XLOOKUP / MIN / MAX / OR / IFERROR 八个函数、
四则与比较运算、百分号字面量、字符串(含 CJK)、跨表引用(未加引号的
CJK 表名)、绝对/相对引用、矩形区间与整列区间。

AST 设计为可 JSON 序列化的 dataclass,数字保留原始文本以便无损回写;
to_formula() 逐节点回写公式文本,与原文逐字符比对构成回归验收。

已知偏离:一元负号与 ^ 的优先级按常规语言处理(Excel 中 -2^2=4),
上游语料无 ^,不影响正确性。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional, Union


# ---------------------------------------------------------------- AST 节点

@dataclass
class Num:
    text: str  # 原始字面量文本,如 "1.5" / "132" / "0.06"

    @property
    def value(self) -> float:
        return float(self.text)

    def to_formula(self) -> str:
        return self.text


@dataclass
class Str:
    value: str

    def to_formula(self) -> str:
        return '"' + self.value.replace('"', '""') + '"'


@dataclass
class Bool:
    value: bool

    def to_formula(self) -> str:
        return "TRUE" if self.value else "FALSE"


@dataclass
class CellPart:
    """区间的一端:列、行至少有其一。col=None 表示整行端,row=None 表示整列端。"""
    col: Optional[str] = None
    row: Optional[int] = None
    col_abs: bool = False
    row_abs: bool = False

    def to_formula(self) -> str:
        s = ""
        if self.col is not None:
            s += ("$" if self.col_abs else "") + self.col
        if self.row is not None:
            s += ("$" if self.row_abs else "") + str(self.row)
        return s


@dataclass
class Ref:
    start: CellPart
    end: Optional[CellPart] = None  # None = 单格引用
    sheet: Optional[str] = None     # None = 本表

    def to_formula(self) -> str:
        s = ""
        if self.sheet is not None:
            name = self.sheet
            if re.fullmatch(r"[\w.一-鿿]+", name):
                s += name + "!"
            else:
                s += "'" + name.replace("'", "''") + "'!"
        s += self.start.to_formula()
        if self.end is not None:
            s += ":" + self.end.to_formula()
        return s


@dataclass
class Call:
    name: str  # 保留原始拼写,含 _xlfn. 前缀
    args: list = field(default_factory=list)

    def to_formula(self) -> str:
        return self.name + "(" + ",".join(a.to_formula() for a in self.args) + ")"


@dataclass
class Bin:
    op: str
    left: "Node"
    right: "Node"

    def to_formula(self) -> str:
        return self.left.to_formula() + self.op + self.right.to_formula()


@dataclass
class Un:
    op: str
    operand: "Node"

    def to_formula(self) -> str:
        return self.op + self.operand.to_formula()


@dataclass
class Pct:
    operand: "Node"

    def to_formula(self) -> str:
        return self.operand.to_formula() + "%"


@dataclass
class Paren:
    inner: "Node"

    def to_formula(self) -> str:
        return "(" + self.inner.to_formula() + ")"


Node = Union[Num, Str, Bool, Ref, Call, Bin, Un, Pct, Paren]


# ---------------------------------------------------------------- 词法

_TOKEN_RE = re.compile(
    r"""
      (?P<WS>\s+)
    | (?P<STRING>"(?:[^"]|"")*")
    | (?P<SHEETQ>'(?:[^']|'')+'!)
    | (?P<SHEET>[\w.一-鿿]+!)
    | (?P<NUM>\d+\.\d+(?:[Ee][+-]?\d+)?|\d+(?:[Ee][+-]?\d+)?|\.\d+)
    | (?P<CELL>\$?[A-Za-z]{1,3}\$?\d+(?![\w.一-鿿]))
    | (?P<COLABS>\$[A-Za-z]{1,3}(?![\w.一-鿿]))
    | (?P<IDENT>[A-Za-z_一-鿿][\w.一-鿿]*)
    | (?P<OP><>|<=|>=|[=<>+\-*/^&%(),:])
    """,
    re.VERBOSE,
)

_CELL_SPLIT_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")
_PURE_COL_RE = re.compile(r"[A-Za-z]{1,3}")


@dataclass
class Token:
    kind: str
    text: str
    pos: int


class FormulaError(ValueError):
    pass


def tokenize(src: str) -> list[Token]:
    tokens: list[Token] = []
    i = 0
    while i < len(src):
        m = _TOKEN_RE.match(src, i)
        if not m:
            raise FormulaError(f"无法识别的字符 {src[i]!r} @ {i}: {src!r}")
        kind = m.lastgroup
        text = m.group()
        if kind != "WS":
            tokens.append(Token(kind, text, i))
        i = m.end()
    return tokens


# ---------------------------------------------------------------- 语法

class Parser:
    def __init__(self, src: str):
        self.src = src
        self.tokens = tokenize(src)
        self.i = 0

    # -- token 游标

    def peek(self, offset: int = 0) -> Optional[Token]:
        j = self.i + offset
        return self.tokens[j] if j < len(self.tokens) else None

    def next(self) -> Token:
        t = self.peek()
        if t is None:
            raise FormulaError(f"公式意外结束: {self.src!r}")
        self.i += 1
        return t

    def expect(self, text: str) -> Token:
        t = self.next()
        if t.text != text:
            raise FormulaError(f"期望 {text!r} 得到 {t.text!r} @ {t.pos}: {self.src!r}")
        return t

    # -- 入口

    def parse(self) -> Node:
        node = self.parse_cmp()
        if self.peek() is not None:
            t = self.peek()
            raise FormulaError(f"多余的 {t.text!r} @ {t.pos}: {self.src!r}")
        return node

    # -- 优先级层次(低 → 高)

    def parse_cmp(self) -> Node:
        node = self.parse_concat()
        while (t := self.peek()) and t.text in ("=", "<>", "<", ">", "<=", ">="):
            self.next()
            node = Bin(t.text, node, self.parse_concat())
        return node

    def parse_concat(self) -> Node:
        node = self.parse_add()
        while (t := self.peek()) and t.text == "&":
            self.next()
            node = Bin("&", node, self.parse_add())
        return node

    def parse_add(self) -> Node:
        node = self.parse_mul()
        while (t := self.peek()) and t.text in ("+", "-"):
            self.next()
            node = Bin(t.text, node, self.parse_mul())
        return node

    def parse_mul(self) -> Node:
        node = self.parse_pow()
        while (t := self.peek()) and t.text in ("*", "/"):
            self.next()
            node = Bin(t.text, node, self.parse_pow())
        return node

    def parse_pow(self) -> Node:
        node = self.parse_unary()
        while (t := self.peek()) and t.text == "^":
            self.next()
            node = Bin("^", node, self.parse_unary())
        return node

    def parse_unary(self) -> Node:
        t = self.peek()
        if t and t.text in ("-", "+"):
            self.next()
            return Un(t.text, self.parse_unary())
        return self.parse_postfix()

    def parse_postfix(self) -> Node:
        node = self.parse_atom()
        while (t := self.peek()) and t.text == "%":
            self.next()
            node = Pct(node)
        return node

    # -- 原子

    def parse_atom(self) -> Node:
        t = self.peek()
        if t is None:
            raise FormulaError(f"公式意外结束: {self.src!r}")

        if t.text == "(":
            self.next()
            inner = self.parse_cmp()
            self.expect(")")
            return Paren(inner)

        if t.kind == "NUM":
            self.next()
            return Num(t.text)

        if t.kind == "STRING":
            self.next()
            return Str(t.text[1:-1].replace('""', '"'))

        if t.kind in ("SHEET", "SHEETQ"):
            self.next()
            if t.kind == "SHEET":
                sheet = t.text[:-1]
            else:
                sheet = t.text[1:-2].replace("''", "'")
            return self.parse_ref(sheet)

        if t.kind in ("CELL", "COLABS"):
            return self.parse_ref(None)

        if t.kind == "IDENT":
            nxt = self.peek(1)
            if nxt and nxt.text == "(":
                self.next()
                self.next()
                args = []
                if self.peek() and self.peek().text != ")":
                    args.append(self.parse_cmp())
                    while self.peek() and self.peek().text == ",":
                        self.next()
                        args.append(self.parse_cmp())
                self.expect(")")
                return Call(t.text, args)
            if t.text.upper() in ("TRUE", "FALSE"):
                self.next()
                return Bool(t.text.upper() == "TRUE")
            # 裸列名(如 SUM(L:L) 的 L)只在区间上下文里合法
            if _PURE_COL_RE.fullmatch(t.text) and nxt and nxt.text == ":":
                return self.parse_ref(None)
            raise FormulaError(f"无法解析的标识符 {t.text!r} @ {t.pos}: {self.src!r}")

        raise FormulaError(f"意外的 {t.text!r} @ {t.pos}: {self.src!r}")

    def parse_ref(self, sheet: Optional[str]) -> Node:
        start = self.parse_ref_part()
        end = None
        if (t := self.peek()) and t.text == ":":
            self.next()
            end = self.parse_ref_part()
        if end is None and start.row is None:
            raise FormulaError(f"孤立的列引用 @ {self.src!r}")
        return Ref(start, end, sheet)

    def parse_ref_part(self) -> CellPart:
        t = self.next()
        if t.kind == "CELL":
            m = _CELL_SPLIT_RE.fullmatch(t.text)
            return CellPart(
                col=m.group(2).upper(), row=int(m.group(4)),
                col_abs=m.group(1) == "$", row_abs=m.group(3) == "$",
            )
        if t.kind == "COLABS":
            return CellPart(col=t.text[1:].upper(), col_abs=True)
        if t.kind == "IDENT" and _PURE_COL_RE.fullmatch(t.text):
            return CellPart(col=t.text.upper())
        raise FormulaError(f"非法的引用端 {t.text!r} @ {t.pos}: {self.src!r}")


def parse_formula(src: str) -> Node:
    """解析一条公式。src 可带或不带开头的 '='。"""
    if src.startswith("="):
        src = src[1:]
    return Parser(src).parse()


# ---------------------------------------------------------------- 自检

if __name__ == "__main__":
    samples = [
        '=58.4+IF($E$24="征人归",5.1,0)',
        '=IF($E$22="三穷致知",2.5%,0)',
        "=B16+1.5%+IF($D$16=\"√\",8%,0)+IF($E$24=\"征人归\",8%,0)",
        "=(B2+B23)*IF($G$2=\"撼天\",1.06,1)",
        "=MAX(B2+B23,C2+B24)*IF($G$2=\"撼天\",1.06,1)",
        "=$L25*VLOOKUP($R25,武学奇术!$A:$ZZ,31,FALSE)",
        "=$G$3+IFERROR(VLOOKUP($AM25,$F$4:$G$7,2,FALSE),0)"
        "+VLOOKUP($R25,武学奇术!$A:$ZZ,2,FALSE)"
        "+SUM(_xlfn.XLOOKUP($S25:$AL25,增益!$A:$A,增益!$B:$B,0))",
        "=SUM(L:L)",
        "=I10/I8",
        "=RD!I14",
        "=I12/143677.47",
        "=132%*1.12*1.7*(1+期望!$B$22/10)",
        "=MIN(IF(VLOOKUP($R25,武学奇术!$A:$ZZ,23,FALSE)=1,100%,$B$8),100%)",
        "=620*(1-IF(F17=\"√\",0.06,0)-IF(D20=\"√\",0.04,0))",
    ]
    for s in samples:
        ast = parse_formula(s)
        out = "=" + ast.to_formula()
        status = "OK " if out == s else "DIFF"
        print(f"{status} {s}")
        if out != s:
            print(f"     -> {out}")
