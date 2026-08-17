# Temper

《燕云十六声》配装 / 毕业率计算工具。纯静态网页,无后端,所有计算在浏览器本地完成。

**在线使用:https://temper.cassianflorin.com**(备用 https://temper-93s.pages.dev)

命名取自 equal temperament(平均律)与金工「回火」双关 —— 游戏装备系统整套是音律隐喻:调律、定音、承音、转律、叠音。

## 它做什么

选流派 → 把游戏面板数值抄进表单 → 实时得到 毕业率 / ADPS / RDPS(无拐)/ 总伤。数值口径与社区竞速计算器完全一致(见下),页面常驻显示数据版本与抓取时间。

## 它是怎么工作的

**不手写伤害公式。** 上游是社区维护的 Excel 计算器(一个流派一个 xlsx,数值绑定竞速轴),Temper 把它们机械地转译成浏览器可执行的 JS:

```
上游 xlsx(h9dh.cn,社区维护)
  │  scripts/sync_upstream.py        清单比对,变化才下载
  ▼
公式解析(scripts/excel_codegen/formula_parser.py)
  │  全语料 63,886 条唯一公式,解析→回写 100% 逐字符往返
  ▼
依赖图 + 拓扑排序(workbook_graph.py)     零环,求值深度 10-11 层
  ▼
JS 生成 + 优化(emit_js.py / optimize_js.py)
  │  静态键查表在生成期折叠为直接引用,整簿重算 ~10ms
  ▼
按标签定位输入/输出格(scan_inputs.py / build_runtime.py)
  │  坐标跨流派漂移,标签词汇稳定;每流派 37-42 个语义字段
  ▼
dist/flows/*.mjs   每流派一个 ESM:set(字段,值) / evaluate()
```

正确性由四道机械验收保证,任何一道红灯都会阻止发布:

1. **解析往返**:全部公式 解析→回写 与原文逐字符比对(`parse_corpus.py`)
2. **缓存值金标准**:生成 JS 重算 126,642 个公式格,与 Excel 自己算的缓存值比对,零失败(`verify_golden.py`)
3. **优化等价性**:随机扰动输入后,优化版与直译版逐格位级一致(`verify_equiv.py`)
4. **接线检查**:字段注入、RD 联动、reset 还原的端到端验证(`verify_runtime.py`)

## 数据来源与致谢

全部数值与竞速轴来自 **[燕云小助手 h9dh.cn](https://www.h9dh.cn/)** 分发的社区计算器,作者 **BiliBili@片雲**。Temper 只做转译与网页化,公式的维护权始终在社区表格里;上游更新版本后,管线重跑即可跟上。

出于对社区作者分发权的尊重,**xlsx 原件不进本仓库**(`upstream/` 在 .gitignore 中),CI 按 `data/upstream-manifest.json` 记录的版本现拉。同步频率为每天两次、无变化时仅请求一次几 KB 的清单页。

## 自动化

`.github/workflows/sync-upstream.yml`,每天北京时间 09:17 / 21:17:

- 抓上游文件清单比对版本,无变化即退出
- 有新表 → 下载 → 上述四道验收 + 输入格摘要(新版式出现未认领输入格即失败)
- 全绿 → 提交 manifest → 构建产物经 artifact 传给 deploy 作业 → 发布 Cloudflare Pages
- 有红 → 不发布,自动开 `upstream-sync-failure` issue(通常意味着上游改了版式,需人工补扫描规则)

## 本地开发

```bash
python3 -m venv .venv && .venv/bin/pip install -r scripts/excel_codegen/requirements.txt
.venv/bin/python scripts/sync_upstream.py            # 拉上游 xlsx 到 upstream/
cd scripts/excel_codegen
../../.venv/bin/python parse_corpus.py               # 验收 1
../../.venv/bin/python verify_golden.py              # 验收 2(需 node)
../../.venv/bin/python verify_equiv.py               # 验收 3
../../.venv/bin/python build_runtime.py              # 产出 dist/flows/*.mjs
../../.venv/bin/python verify_runtime.py             # 验收 4
cd ../.. && python3 scripts/build_site.py            # 组装 dist/site
python3 -m http.server 8797 -d dist/site             # 本地预览
```

## 仓库结构

```
web/                      前端源码(纯静态,无构建链)
scripts/sync_upstream.py  上游同步
scripts/build_site.py     站点组装
scripts/excel_codegen/    codegen 管线(解析/依赖图/生成/优化/扫描/构建/验收)
data/constants.json       110 阶满值表、流派定义(全部经上游 xlsx 逐项复核)
data/flow-inputs.json     每流派 语义字段 -> 单元格 映射(扫描生成)
data/upstream-manifest.json  上游文件版本清单(CI 维护)
docs/xlsx-anatomy.md      上游 xlsx 结构解剖(管线设计的依据)
```

## 边界与路线

- **配装求解器与 OCR 录入尚未实现**:两者都依赖「词条 → 面板」的 110 阶五维换算系数,流传数据是 96 版本的,未经验证不能用(见 `data/constants.json` 的 `knownGaps`)。数据补齐前不编造。
- 牵丝霖的计算器是上游有意设计的 15 秒爆发轴,其毕业率仅对本流派有意义,不可与其他流派横向比较。
- 词条满值与承音系数(×0.94)已对上游 DIY 计算器逐项复核。

## 免责

本项目与网易及《燕云十六声》官方无关;数值为社区测算结果,随版本与社区表格更新而变化,不构成官方数据。
