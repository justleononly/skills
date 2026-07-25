---
name: ip-sales-analysis
description: IP销售分析报告生成 - 基于Excel销售数据，按照「新IP分析方法」标准分析框架生成结构化的Word分析报告(.docx)。使用场景：(1) IP销售活动结束后需要输出标准化分析报告，(2) 需要按固定分析框架处理销售数据，(3) 需要生成包含图表、表格、高亮结论的专业文档。
---

# IP销售分析

## 使用流程

1. 确认用户提供了源数据Excel文件路径
   > **重要原则：** 数据预处理阶段不应对任何字段为"未知"或null的行做全局过滤。所有过滤逻辑由 `generate_report.py` 在对应分析维度内部自动处理：
   > - 性别分析时自动过滤 `未知`（仅统计 `男`/`女`）
   > - 年龄分析时自动过滤 NaN（通过 `dropna()`）
   > - 性别&年龄交叉分析时同时过滤性别未知和年龄NaN
2. 如数据中包含 `生日` 字段但无 `年龄` 字段，需先执行数据预处理（计算年龄）
3. 如数据中的时间列名称不是标准的 `下单时间`（如为 `行为时间`），需在预处理中重命名
4. 如数据中包含首次活跃时间字段（如 `抽卡机首次访问时间（虚拟属性）`），可重命名为 `注册时间` 以支持新老用户分析
   > **新老用户逻辑：** 逐行比较`注册时间`（首次活跃时间）与`下单时间`，若首次活跃时间距下单时间不超过30天则为新用户，否则为老用户。支持月度范围格式（如 `2026-07-01~2026-08-01`），取范围开始日期进行对比
5. 使用 `scripts/generate_report.py` 生成分析报告
6. 使用 `scripts/validate_report.py` 校验报告结构

## 核心脚本

### generate_report.py — 报告生成器

Python脚本，读取Excel销售数据，计算指标，生成图表，输出Word报告。

**必需参数：**
- `--excel <路径>` : 销售数据Excel文件路径

**可选参数：**
- `--output <目录>` : 报告输出目录（默认当前目录）

**用法：**
```bash
PY=/Users/lhy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
$PY scripts/generate_report.py --excel /path/to/data.xlsx --output /path/to/output
```

### validate_report.py — 结构校验器

校验生成的.docx报告与「新IP分析方法」模板的结构一致性。

**用法：**
```bash
PY=/Users/lhy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
$PY scripts/validate_report.py --report /path/to/report.docx
```

## 分析框架

报告严格遵循以下分析框架（来自「新IP分析方法」模板）：

### 报告章节结构

1. **关键结论** — 数据驱动的核心发现提炼
2. **基础销售数据** — GMV、付费用户数、客单价、复购率
3. **销售数据趋势** — 日销售额折线图 + 趋势解读
4. **付费用户画像**
   - 4.1 性别分布（饼图：用户数 + 销售额）
   - 4.2 年龄分布（饼图：用户数 + 销售额，需有 `年龄` 或 `生日` 字段）
   - 4.3 性别&年龄交叉分布（表格按销售额降序排列，需有年龄字段）
   - 4.4 会员等级分布（饼图：用户数 + 销售额，需有 `会员等级（标签）` 字段）
   - 4.5 新老用户分布（饼图：用户数 + 销售额，需有 `注册时间` 字段）
5. **分析总结与建议**
6. **指标计算说明** — 各指标的定义与计算逻辑

> 模板结构不可删除，但可根据数据情况新增分析维度和结论。

### 核心指标定义

| 指标 | 定义 |
|------|------|
| 总销售额(GMV) | 所有有销售记录行的销售额加总 |
| 付费用户数 | 有销售额>0的唯一member_code数量 |
| 客单价(AOV) | 总销售额 / 付费用户数 |
| 复购用户 | 付费次数≥2的唯一member_code |
| 复购用户占比 | 复购用户数 / 付费用户数 × 100% |
| 复购销售额占比 | 复购用户的销售额之和 / 总销售额 × 100% |
| 新用户 | 逐条记录比较首次活跃时间与其下单时间，若首次活跃时间距下单时间不超过30天则视为新用户 |
| 老用户 | 首次活跃时间距下单时间超过30天 |

### 数据清洗规则

- `销售额` 列的 `'null'` 字符串 或 NaN → **删除**该行（不保留、不转为0）
- `下单时间` 列为 `合计` 的行 → 自动过滤（此为结构汇总行，非数据行）
- `性别`、`年龄` 等字段的 `未知` 或 null 值 **不** 在预处理阶段过滤，而是留到对应分析维度内部按需排除
  - 脚本的性别分析会自动过滤 `未知`（仅统计 `男`/`女`）
  - 年龄分析自动通过 `dropna()` 过滤 NaN
  - 交叉分析同时过滤性别未知和年龄NaN
  - 此举保证各维度分析独立、数据最大化可用

### 图表输出

- **日销售额趋势折线图**（带数值标签，红色曲线+填充）
- **性别分布饼图**（左右双饼图：用户数 + 销售额，标签在外围显示"名称+占比"，黑色字体）
- **年龄分布饼图**（左右双饼图：用户数 + 销售额，标签在外围显示"年龄段+占比"，黑色字体）
- **会员等级分布饼图**（左右双饼图：用户数 + 销售额，标签在外围显示"等级+占比"，黑色字体）
- **新老用户分布饼图**（左右双饼图：用户数 + 销售额，标签在外围显示"用户类型+占比"，黑色字体）

所有饼图使用 `labeldistance=1.15` 将标签放置在饼图外围，黑色文字（`color="black"`），显示名称和百分比，不显示具体数值。图表输出在报告输出目录（与docx同目录），会自动嵌入到docx中。

**字体说明：** 报告的文本元素（标题、正文、表格、高亮框、元数据区）统一使用 `Microsoft YaHei`（微软雅黑）字体。脚本自动适配中文字体。如遇图表中文显示异常，需确认matplotlib字体缓存路径（设置 `MPLCONFIGDIR=/tmp/matplotlib` 可解决权限问题）。

**MPLCONFIGDIR使用示例：**
```bash
MPLCONFIGDIR=/tmp/matplotlib $PY scripts/generate_report.py --excel /path/to/data.xlsx --output /path/to/output
```

## 数据字段适配

脚本自动检测Excel中的列名来决定可分析维度：

- **必须有**：`下单时间`、`member_code`、`销售额`
- **如有则分析**：`性别`、`会员等级（标签）`、`付费次数`
- **年龄分析**：`年龄`（直接作为整数年龄）或 `生日`（格式 `YYYY-MM-DD`，脚本自动计算基于分析周期中点的年龄）
- **新老用户分析**：`注册时间`（首次活跃时间，脚本逐行对比`注册时间`与`下单时间`，若不超过30天则视为新用户）
- **建议补充**：`注册时间`（支持更完整的画像分析，如新老用户对比）

> **注意：** 如果原始数据中的时间列名是 `行为时间` 而非 `下单时间`，需要在预处理阶段重命名。如果首次活跃时间字段名为 `抽卡机首次访问时间（虚拟属性）`，也建议在预处理中重命名为 `注册时间`。

如果数据缺少某些字段，报告对应章节会显示友好提示，不影响其他部分生成。

### 数据预处理（列名适配 + 生日→年龄）

如原始数据包含 `生日` 字段但无 `年龄` 字段，或列名不标准，需要在运行报告生成前进行预处理。示例：

```bash
PY=/Users/lhy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
$PY -c "
import pandas as pd
from datetime import datetime

df = pd.read_excel('数据文件.xlsx')

# 列名适配
df = df.rename(columns={
    '行为时间': '下单时间',                    # 如有需要
    '抽卡机首次访问时间（虚拟属性）': '注册时间',  # 如有需要
})
# 注册时间解析：首次活跃时间可能是月范围格式（如 2026-07-01~2026-08-01），取其开始日期

# 去掉销售额为null的行（字符串'null'或NaN均视为无效）
def is_sales_null(val):
    if pd.isna(val): return True
    s = str(val).strip()
    return s.lower() == 'null' or s == ''
df = df[~df['销售额'].apply(is_sales_null)].copy()
df['销售额'] = df['销售额'].astype(float)

# 去掉行为时间=合计的结构汇总行
df = df[df['行为时间'] != '合计'].copy()

# 生日→年龄
ref_date = datetime(2026, 7, 7)  # 分析周期中点
def compute_age(birth_str):
    if pd.isna(birth_str): return None
    birth = datetime.strptime(str(birth_str).strip(), '%Y-%m-%d')
    age = ref_date.year - birth.year
    if (ref_date.month, ref_date.day) < (birth.month, birth.day):
        age -= 1
    return age

df['年龄'] = df['生日'].apply(compute_age)
df.to_excel('数据文件_处理后.xlsx', index=False)
"
```

预处理后，生成的Excel文件即可被 `generate_report.py` 完整分析年龄相关章节和各维度。

> **处理原则：** 预处理仅做列名重命名和类型转换，**不要过滤任何字段为"未知"或null的行**。脚本中的分析函数会自动处理各维度的过滤：
> - `性别` 为 `未知` 的行在性别分析时会被排除，但仍参与年龄、会员等级等其他维度分析
> - `年龄` 为 `未知`/NaN 的行在年龄分析时会被排除，但仍参与性别等其他分析
> - 新老用户：逐行比较`注册时间`与`下单时间`，距下单时间不超过30天视为新用户，否则为老用户
> - 此举保证各维度分析独立、数据最大化可用

### 数据预处理（strict XLSX格式兼容）

某些Excel文件使用微软的 `conformance="strict"` 格式，导致pandas/openpyxl无法直接读取。此时需先通过解析XML结构提取数据：

```bash
PY=/Users/lhy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
$PY -c "
import zipfile, pandas as pd
from xml.etree import ElementTree as ET

fpath = '数据文件.xlsx'
zf = zipfile.ZipFile(fpath)
ns = {'s': 'http://purl.oclc.org/ooxml/spreadsheetml/main'}
ss_xml = zf.read('xl/sharedStrings.xml')
ss_tree = ET.fromstring(ss_xml)
shared_strings = [si.find('s:t', ns).text if si.find('s:t', ns) is not None else '' for si in ss_tree.findall('.//s:si', ns)]
ws_xml = zf.read('xl/worksheets/sheet1.xml')
ws_tree = ET.fromstring(ws_xml)
rows = ws_tree.findall('.//s:row', ns)
data = []
for row_elem in rows:
    cells = row_elem.findall('s:c', ns)
    row_data = {}
    for cell in cells:
        col_letter = ''.join(c for c in cell.get('r') if c.isalpha())
        v = cell.find('s:v', ns)
        t = cell.get('t', '')
        if v is not None and v.text:
            row_data[col_letter] = shared_strings[int(v.text)] if t == 's' else v.text
        else:
            row_data[col_letter] = None
    data.append(row_data)

df = pd.DataFrame(data)
col_names = list(df.iloc[0])
df = df.iloc[1:].reset_index(drop=True)
df.columns = col_names
df.to_excel('数据文件_处理后.xlsx', index=False)
"
```

预处理后，后续步骤与普通Excel文件一致。

## 结果校验

每次生成报告后使用 `validate_report.py` 做结构校验：

```bash
$PY scripts/validate_report.py --report "IP销售分析报告_2026-07-03_2026-07-08.docx"
```

校验通过后，手动检查：
1. 图表是否正确嵌入、尺寸合理
2. 饼图标签是否在外围显示清晰、不重叠（黑色文字，名称+占比）
3. 结论高亮框是否醒目（黄色底色）
4. 文字格式是否统一（微软雅黑，字号、颜色、对齐）
5. 表格数据是否与Excel源数据一致

## 可扩展性

此脚本设计为可扩展的分析框架。如需新增分析维度：

1. 在 `compute_metrics()` 中计算新指标
2. 在 `generate_report()` 中添加对应章节
3. 添加新的图表函数（如有需要），注意饼图标签统一使用 `labeldistance=1.15` 和 `color="black"`
4. 更新 `EXPECTED_STRUCTURE` 列表

## 依赖环境

使用Codex捆绑的Python运行时，位于：
```
/Users/lhy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
```
依赖已预装：pandas, openpyxl, matplotlib, python-docx, pdfplumber, numpy, Pillow
