# 克隆与合成（Clone & Compose）

> **这个文件解决什么问题**：把已经确认的 `copy.md` + `visual-plan.md` 变成可执行的 `build-plan.json` 与真实 `.pptx`——克隆哪个样张、改哪几个 `shapeId`、删哪些装饰、缺角色时怎么用四步升级补出来，以及所有实测踩过的坑。

## 先决策：这一页走哪条路

```
模板能"套版式"吗？
├─ add_slide(layout) 出来有标题栏/logo → A 版式复用（填占位符）
└─ 出来是白页（中文高校模板主流）
   ├─ 有同角色的可用样张 → B 样张克隆（主力路径）
   ├─ 只有近似角色（3 栏 / 白底页 / 图文页）→ B' 克隆 + 改造（删栏、换槽位、加元素）
   └─ 一个都不匹配 → C 风格合成（保视觉 token，正文区放原生图表/表格/文本）
```

**任何一条路都不允许改内容去迁就版式。** 放不下 → 回 Step 2 拆页或改结构，并在确认门 1 里重新确认；生成阶段偷偷合并板块 = 不合格。判定阈值（占位符数、浮动文本框数、装饰图占比、10 秒白页验证）见 `template-profile.md` 的"三条路径判定表"。

## 样张克隆：六个步骤，一步都不能少

```python
# 1. 基于源页的版式新建一页（版式是空壳，只为拿到合法的 slide part）
new = prs.slides.add_slide(src.slide_layout)
for shp in list(new.shapes):                    # 2. 清掉 add_slide 自带的占位符
    shp._element.getparent().remove(shp._element)

# 3. 复制背景（p:bg 不在 spTree 里，单独拷）
bg = src._element.find(qn("p:cSld")).find(qn("p:bg"))
new_csld.insert(0, deepcopy(bg))                # 先移除新页原有 p:bg

# 4. 深拷贝形状树：跳过 nvGrpSpPr / grpSpPr（新页已有，重复会损坏文件）
for child in src.shapes._spTree:
    if etree.QName(child).localname in ("nvGrpSpPr", "grpSpPr"):
        continue
    new.shapes._spTree.append(deepcopy(child))

# 5. 关系重映射：图片/超链接的 r:embed、r:id 还是源页的 rId，必须在新 part 上重建
for el in new._element.iter():
    for attr in list(el.attrib):  # list() 避免迭代中改属性
        if attr.startswith("{" + R + "}"):
            rel = src.part.rels.get(el.attrib[attr])
            el.set(attr, new.part.relate_to(rel.target_part, rel.reltype))

# 6. 生成完毕后删除模板原始页（先 drop_rel 再移除 sldId，否则关系残留）
```

**为什么必须第 5 步**：pptx 里的图片是 `r:embed="rId3"` 这种关系 ID，`deepcopy` 只复制了 XML 字符串。不重映射 → 新页的 `rId3` 指向别的东西（或不存在）→ 图片空框、校徽丢失。外部链接走 `relate_to(rel.target_ref, rel.reltype, is_external=True)`；漏处理 `r:id` 会丢超链接与部分图片。

**克隆不继承**备注、动画、批注、chart part → 备注单独写：`new.notes_slide.notes_text_frame.text = page["notes"]`。

## 文本替换：整框替换，不是逐 run 替换

```python
def replace_text(shape, text):
    txbody = shape._element.find(qn("p:txBody"))
    first_p = txbody.find(qn("a:p"))
    ppr = deepcopy(first_p.find(qn("a:pPr")))        # 保留首段段落属性（对齐/行距/项目符号）
    rpr = deepcopy(first_p.find(qn("a:r")).find(qn("a:rPr")))  # 保留首个 run 的字号/字体/颜色（无 run 时判空）
    for p in txbody.findall(qn("a:p")):
        txbody.remove(p)                             # 清空，按 "\n" 重建 a:p 列表
    for line in text.split("\n"):
        p = etree.SubElement(txbody, qn("a:p"))
        p.append(deepcopy(ppr))
        r = etree.SubElement(p, qn("a:r"))
        r.append(deepcopy(rpr))
        etree.SubElement(r, qn("a:t")).text = line
```

| 做法 | 结果 |
|---|---|
| ✅ 整框替换：保留首个 `a:p` 的 `pPr` + 首个 `a:r` 的 `rPr`，重建段落 | 字号/字体/颜色不丢，多行按 `\n` 分段 |
| ❌ 按 run 逐个替换 | 模板 run 数 ≠ 文案行数时文字**重叠乱码**（实测标题糊成一团） |
| ❌ 清空后重新 `shape.text = ...` | 所有 `rPr` 丢失 → 字号回退 18pt / 字体变默认 |
| ❌ 用 `add_run` 追加 | 旧文字仍在，新文字叠在后面 |

## 几何与字号覆盖：先改框，再改字号

plan 每条 `texts[]` 支持可选几何/字号覆盖：

```json
{"shapeId": 2, "text": "研究设计：三维面板 + PPML + 工具变量",
 "x": 0.95, "y": 0.30, "w": 11.30, "h": 0.60, "size": 28}
```

固定优先级：**改文案（减字）→ 改几何（`x/y/w/h` 扩框）→ 改字号（`size`）→ 换样张**。

- 标题换行：先加宽 `w`，不要先缩 `size`；
- 正文溢出：先拆页，再考虑 `size`（正文低于 18pt 就是不合格，不是"将就一下"）；
- 改框会破坏模板对齐 → 改完必须渲染目检（左右边距、与装饰的相对位置）。

## 合成元素（Path B' / Path C 用）

三个 `addShapes[].kind`，字段与 `build_deck.py` 完全一致：

**`table`（3 列 × N 行，清单/对比/符号表）**

| 字段 | 必填 | 说明 |
|---|---|---|
| `data` | ✅ | 二维数组，第 1 行是表头；列数由 `data[0]` 决定 |
| `x/y/w/h` | ✅ | 英寸；表宽建议 = 标题栏宽（约 12.08 in）内缩 0.3 in |
| `colWidths` | 建议 | 每列宽（英寸），**和 = w**，否则右侧留白不均 |
| `rowHeights` | 建议 | 首行 0.65，数据行 0.95（中文 16pt 两行的实测高度） |
| `fontPt` | 建议 | 表格 14–16pt（投影下限，别再小） |
| `headerColor` | 可选 | 表头**字色**（品牌色，如 `FFFFFF` 需深底） |

> ⚠ 实测限制：当前脚本只改表头**字色**，没改表头**底色**——底色仍是 python-pptx 默认的 Office 蓝。在脚本补 `headerFill` 之前，优先用"白底 + 品牌色表头字 + 细线"的观感，或生成后处理 `cell.fill`。垂直居中是脚本内置的（`MSO_ANCHOR.MIDDLE`），不用自己写。

**`chart`（原生图表，不是图片）**

| 字段 | 必填 | 说明 |
|---|---|---|
| `type` | ✅ | `column` / `bar` / `line` / `pie` |
| `categories` / `series` | ✅ | `series: [{name, values}]`，`values` 长度 = `categories` 长度 |
| `colors` | ✅ | **只能取模板 `actualColors` 的 hex**（如西财蓝 `376FFF`、`FFCD69`） |
| `fontPt` | 建议 | 12–14pt；图例 >1 个系列时脚本自动放底部 |
| `title` / `titlePt` | 可选 | 答辩页标题已在标题栏，图表内标题建议**不要**（避免一页两个标题） |

**`text`（图注/注释/数据来源）**

| 字段 | 必填 | 说明 |
|---|---|---|
| `x/y/w/h` | ✅ | 图注放图下方 0.2–0.3 in，高度 0.45–0.5 in |
| `text` | ✅ | `\n` 分段 |
| `size` | 建议 | 12pt（图注下限） |
| `color` | 建议 | 灰蓝，如西财蓝 `637490`；不要纯黑 |
| `bold` / `align` | 可选 | `center` 用于居中图注 |

**图表必须有图注 + 数据来源**（硬规则 10），图注写"数据来源：xxx；样本区间；N=xxx"，方便评委追问。

## 装饰删除原则

| 保留（永久） | 可删（与内容无关时） |
|---|---|
| 标题栏 / 标题下分隔线 / 色带 | 校园照片、风景图、示例人物照 |
| 校徽 / 院徽 / logo（右上角） | 装饰用分栏线（当该页改成表格/图表时） |
| 页脚页码、学校名 | 示例大编号（01/02/03，当编号与内容不对应时） |
| 背景底图（调低透明度后不压字） | 与正文重复的装饰形状 |

判断标准：**删掉它，这页还像不像学校的模板？** 像 → 可删；不像 → 保留。删不准就保留，渲染后看文字有没有压图。

## ⚠ 模板与内容不匹配：四步升级策略

这是本 skill 最容易出错、也最容易被评委看出来的地方。内容项数与模板槽位数不一致时，**按下面四步逐级升级，不许跳步，不许合并/删减内容**。

| 步 | 做什么 | 具体动作（西财蓝实测） | 禁止 |
|---|---|---|---|
| **1. 保留视觉 token** | 把模板的视觉身份固定下来，后面所有改造都围绕它 | 标题栏几何（`x0.64 y0.21 w12.08 h0.57`）、标题下分隔线（`y0.79`）、校徽位（右上 `10.41,0.00,2.62×1.06`）、实际品牌色（`376FFF`/`1271BA`/`637490`/`FFCD69`）、字体（微软雅黑）、左右页边距（`0.64 in`）、圆角卡片风格 | 换字体、换配色、改标题栏位置 |
| **2. 改造页面结构** | 删/并/拆槽位，让结构匹配内容项数 | "4 项内容 vs 模板 3 栏" → 改表格（3 列 × 5 行）；"5 个阶段 vs 模板 1 图 3 段" → 改时间线；"2 方案对比 vs 模板 3 栏" → 保留两栏、第三栏删掉并加宽 | 把 4 项合并成 3 项；删掉一项 |
| **3. 删除无关装饰腾空间** | 只删与内容无关的装饰，给学生腾出正文区 | 删校园照片（`30/31/32`）、分栏线（`24/25/26`）、示例大编号（`5/10/14`）、示例小标题与正文（`3/4/6/7/11/12`） | 删标题栏、校徽、色带、页脚 |
| **4. 合成新页面** | 白底 + 保留的标题栏 + 原生元素 | 克隆同色系页面 → 删装饰 → `add_table`/`add_chart`/`add_text`；尺寸与 token 对齐；颜色取实际用色 | 自创模板外装饰、加阴影/渐变/emoji |

**铁律**：容量不够就"多一页"或"换一个容量更大的版式"，**不许为套模板把不同板块塞进一页**。
**槽位守恒（2026-11 新增，用户明确要求）**：内容项数 ≠ 模板槽位数时，必须改造格式，**不许合并、不许删减**。三种正解：

| 内容 | 模板槽位 | 错误 | 正解 |
|---|---|---|---|
| 2 个假说 | 3 卡片 | 塞"理论基础"凑第 3 张 | 换 2 块布局（如 P9 的两个编号块） |
| 5 类稳健性 | 4 卡片 | 只放 4 个，第 5 个丢进备注 | 合成 5 行表格，5 项全上 |
| 5 个设计要素 | 3 行编号列表 | 把 3 个变量并成 1 行 | 合成 5 行表格 |

交付前用 `scripts/check_plan.py` 做覆盖检查：`copy.md` 的每条内容必须能在对应页的幻灯片里找到；缺一条即返工。这一步的取舍必须回到确认门 1/2 跟用户确认，不能由生成脚本静默决定。

### 真实案例：P6 有 4 个数据源，模板只有 3 栏

- **内容**：企业层 / 财务层 / 供应链 / 东道国层（4 个数据源，对应论文 4.1 节）。
- **模板**：西财蓝第 7 页是白底三栏（每栏"小标题 + 正文 + 一幅照片"）。
- **错法 A（合并）**：把"供应链"和"东道国层"并成一栏 → 信息丢失，评委问"ICRG 具体怎么用"时没有落脚点。
- **错法 B（删项）**：删掉一层 → 内容不完整，等于对评委隐瞒数据来源。
- **正解（四步升级）**：克隆第 7 页 → 删 15 个无关装饰/示例 shape（`3,4,5,6,7,10,11,12,14,24,25,26,30,31,32`）→ 插原生表格 3 列 × 5 行（数据层 / 来源 / 用途），`fontPt 16`、`colWidths [1.4,4.75,5.25]`、`rowHeights [0.65,0.95,0.95,0.95,0.95]`，表格区 `x0.95 y1.4 w11.42 h4.45` → 底部 `add_text` 图注（`size 12`、`color 637490`、`y6.15`）。结果：4 项全部呈现、一眼可比、仍然是西财的视觉。
- 这一页正是**单页样张**（确认门 3）的候选：同时验证"克隆 + 删装饰 + 插原生表格"三件事。

## 实测踩坑清单

### 中文折行丢字（2026-11 新增，典雅红模板实测）

**现象**：克隆的卡片正文在渲染图里"每个折行的首字消失"（如"数字经济"→"数字/济"），但 pptx 里的文本是完整的。
**根因**：模板段落带 `hangingPunct="1"`（标点悬挂）+ `algn="just"`（两端对齐），LibreOffice 渲染中文时会在换行处丢字符。PowerPoint 不一定复现，但无法在交付前验证。
**修法**：`build_deck.py` 的 `replace_text()` 已自动规范化——遇到 `algn="just"` 改 `l`、`hangingPunct="1"` 改 `0`。视觉差异极小，但必须做。
**教训**：结构校验（文本完整）通过 ≠ 渲染正确；**必须逐页看渲染图**，尤其卡片/多栏正文。


| # | 现象 | 原因 | 修法 |
|---|---|---|---|
| 1 | 克隆后图片/校徽丢失、空框 | 只 `deepcopy` 未重映射 `r:embed`/`r:id` | 遍历 XML 里的 `r:*` 属性，用 `relate_to` 重建 |
| 2 | 标题文字重叠、糊成一团 | 按 run 逐个替换 | 整框替换（保留首 `pPr` + 首 `rPr`） |
| 3 | 字号/字体丢失，回退 18pt | 清空 `a:p` 后新 run 没有 `rPr` | 深拷贝首 run 的 `rPr` 到每个新 run |
| 4 | 改一页图表，另一页跟着变 | 克隆含图表的页 → 共享同一 chart part | 重建 chart part；更稳的是 Path C 新插图表 |
| 5 | 文字溢出文本框 | python-pptx **不计算 autofit** | 按容量字数控制 + 几何覆盖 + 拆页 |
| 6 | `fontSize` 是 null，不知道多大 | 字号继承自版式/母版 | 渲染一页实测，或 plan 里显式 `size` |
| 7 | 渲染图里宋体变黑体 | 缺字体静默回退（思源宋体 CN → PingFang SC） | 读 `missingFonts`、按 `assets/font-fallbacks.json` 替换、明确告知用户 |
| 8 | 新页多出空白占位符 | `add_slide` 自动生成占位符 | 克隆前清空新页 shapes |
| 9 | 文件里还留着模板原页 | 只移了 `sldId` 没 `drop_rel` | 先 `drop_rel(rId)` 再移除 `sldId` |
| 10 | 备注全空 | 克隆不继承 notes | `new.notes_slide.notes_text_frame.text = ...` |
| 11 | 渲染 CLI 报输入过大 | 默认输入上限 **64MiB** | 加 `--max-input-bytes 400000000` |
| 12 | 渲染报目录已存在 | `--output-dir` 必须**不存在** | 每次换新目录名，或用时间戳 |
| 13 | 本机好看，答辩机错位 | LibreOffice 与 PowerPoint 的行距/字体/阴影渲染不同 | 终检在 Windows+Office 上打开一次，并备份 PDF |
| 14 | 页面"不像学校模板了" | 删装饰删过头 | 按保留清单逐项核对标题栏/校徽/色带/页脚 |

## 脚本命令速查

```bash
# 1) 模板画像（Step 3，只做一次）
python3 <SKILL_ROOT>/scripts/inspect_template.py "<模板.pptx>" -o out/profile

# 2) 单页样张（确认门 3）：--only 只生成指定页，支持逗号分隔
python3 <SKILL_ROOT>/scripts/build_deck.py "<模板.pptx>" plan.json -o out/样张.pptx --only 4

# 3) 批量生成（确认门 3 通过后才跑）
python3 <SKILL_ROOT>/scripts/build_deck.py "<模板.pptx>" plan.json -o out/答辩.pptx

# 4) 结构校验（P0）
python3 <SKILL_ROOT>/scripts/validate_deck.py out/答辩.pptx --scene bachelor --out out/qa-report.json

# 5) 体积治理（目标 <15MB）
python3 <SKILL_ROOT>/scripts/optimize_pptx.py out/答辩.pptx -o out/答辩-压缩.pptx \
    --max-width 1920 --quality 82 --min-kb 800

# 6) 渲染目检（封装脚本：只渲染改动页 + 拼图 + missingFonts）
python3 <SKILL_ROOT>/scripts/render_check.py out/答辩-压缩.pptx -o out/render-new \
    --pages "4,6" --dpi 110 --montage
#    大模板优先用它（内部按 --max-mb 换算 --max-input-bytes，默认 400）；
#    缺字体显式映射：
#    --font-fallbacks '[["微软雅黑","PingFang SC"],["思源宋体 CN","Songti SC"]]'

# 6b) 直接调 LOKit CLI（与 6 等价；--output-dir 必须不存在）
node <LOKIT_CLI> render --input out/答辩-压缩.pptx --output-dir out/render-new \
    --pages "4,6" --dpi 110 --max-input-bytes 400000000
```

> 生成后**只看改动页**，别全量重渲染；读返回 JSON 的 `missingFonts`，缺字体要当面告知用户。

## 反例

- ❌ **克隆完直接改文字不看渲染**：图片丢了、标题换行了都不知道 → 单页样张的意义就是先撞这一堵墙。
- ❌ **为了让 4 项塞进 3 栏把正文缩到 12pt**：违反信息预算，且模板 12pt 本就不是交付状态。
- ❌ **删掉校徽"让页面更干净"**：违反硬规则 5（忠于模板），一眼就能看出不是学校的模板。
- ❌ **克隆图表页后改数据**：chart part 共享 → 所有克隆页一起变，且容易改出与论文不一致的数字。
