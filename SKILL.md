---
name: physics-lab-report
description: 把大学物理实验的原始材料（手写照片、空白报告模板、实验数据）做成一份可直接提交的实验报告 PDF：照片页打头、写数据处理与结论、画图、导出并按实验名称命名。当用户提供手写照片 / 空白 docx / 实验数据，并要求做实验报告、填数据、处理数据、画曲线、修订报告或转 PDF 时使用。不适用于没有实验数据的普通写作，也不适用于纯排版。
---

# 大学物理实验报告

把「原始材料 → 可提交的 PDF」这条流水线跑完。

## 三条总原则

**① 照片页打头，不要"翻译"成打字版。**
用户交来的是他手写的原件（预习页、数据记录页，有时还有第三页）。报告**最前面那几页就是这些照片本身** —— 一张照片一页，页数跟照片张数走（常见 2 张，也可能 3 张以上），不是把内容重新打一遍。见「流程 1」。

**② 能不写就不写。**
一句话加不加，如果不影响题目要求、数据结果、可复现性或必要解释，就删掉。适用于正文、标题、图表说明和最终回复。每条事实只说一次。

**③ 绝不编造数据。**
手写看不清就重新切片放大；仍看不清就**留空并告诉用户**，不要猜。

❗报告正文里**不要**出现「若采用另一个数值则…」「这说明读数误差较大」这类自我说明——用户明确说过这种话"一眼就暴露是 AI 生成"。原始数据的问题**只在对话里说**，由用户拍板。

## 输入 / 输出

| | 内容 |
|---|---|
| **输入** | 手写照片（预习页、数据记录页，可能 3 张以上）、原始空白 docx、有时附口头要求 |
| **输出** | **`<实验名称>.pdf`** —— 文件名就是文档里的实验名称 |

❌ 不要 `最终版` / `v10` / `final` 之类后缀；❌ 不要留一堆同义文件让用户猜哪个是最终的。

## 流程

### 0. 建工作目录
`<工作区>/exp-report<实验号>\` 下放：`original.docx`（模板）、`page1.jpg` / `page2.jpg` / `page3.jpg`…（照片，张数不定）。

### 1. 定版式 —— 照片页打头，一张照片一页 ⚠️

报告最前面那几页是**用户给的照片原图**，一张照片一页，页数跟照片张数一致（**常见 2 张，也可能 3 张或更多 —— 不要假设"就两页"**，开跑前先 `Get-ChildItem page*.jpg` 数一下）：

| 页 | 内容 | 谁提供 |
|---|---|---|
| 1 ~ N | `page1.jpg`、`page2.jpg`、`page3.jpg`…（预习页、数据记录页，可能还有第三页） | 照片，一张一页 |
| N+1 起 | 数据处理、曲线图、结论及现象分析、讨论题 | **你写** |

做法：**删掉模板里「数据处理」标题之前的全部内容**（预习区、空数据表、签字框全部作废），再把所有照片放到文档最前、每张占一页。

```powershell
# 照片按文件名排序（page1、page2、page3…），张数不固定
$photos = @(Get-ChildItem "$work\page*.jpg" | Sort-Object Name)
if ($photos.Count -lt 1) { throw "没找到 page*.jpg 照片" }

# ① 删掉「数据处理」之前的一切
$s = (GetPara $doc "数据处理").Range.Start
if ($s -gt 1) { $doc.Range(0, $s).Delete() }

# ② 第一张放文档最前
$pic1 = $doc.InlineShapes.AddPicture($photos[0].FullName, $false, $true, $doc.Range(0,0))
$pic1.LockAspectRatio = $true; $pic1.Width = $availW
$doc.Range($pic1.Range.End, $pic1.Range.End).InsertParagraphAfter()

# ③ 其余照片「倒序」插到「数据处理」标题前（倒序才不会把顺序插反）
for ($i = $photos.Count - 1; $i -ge 1; $i--) {
    $tPara = GetPara $doc "数据处理"
    $tPara.Range.InsertParagraphBefore()
    $holder = $tPara.Previous()
    $p = $doc.InlineShapes.AddPicture($photos[$i].FullName, $false, $true, $holder.Range)
    $p.LockAspectRatio = $true; $p.Width = $availW
    $holder.PageBreakBefore = $true
}
```

- ❗**不要把手写照片"翻译"成打字版**。重新打一遍等于替用户重做作业，还会丢掉他的笔迹、涂改、签名。
- 只有用户**明确说**「帮我把答案打进去」时才走打字路线。
- **整页扫描是最合适的形态**（自带页眉页脚、签名、涂改痕迹），照原样插入即可，不要裁、不要美化。局部特写照只能当插图，不能当页。
- 照片插进去会和文档页眉形成**双页眉**（照片自带一条）——这是历史报告的既有样式，**不用管**。

### 2. 读手写数据 —— 只为了算和画，不是为了录入

1. 先整张 `read_image` 看布局，判断表格结构
2. 用 sharp **按行切窄条**（宽 300~600px）再放大 3~8 倍，逐块 `read_image` 精读
3. **不要用 `vision_glance` 读表格** —— 它经常只给前 3 行就截断，或输出推理而非数字
4. **用公式反查**：列出该实验的公式，看哪几个读数必须互相吻合，用自洽链锁定模糊字迹

> 例（光速实验）：`l = X1 − X2` → 1.682 − **0.209** = 1.473，据此锁定「0.209」而不是 0.239
> 例（光电效应）：三张表的 ν 列完全相同，Uc 随孔径的变化必须单调可解释

⚠️ 自洽链不是万能的——学生自己算错的数也可能"自洽"。用户指出数据错了，一律按用户说的改。

### 3. 算结果

- **题目要求几种方法就做几种**（说"逐差法和最小二乘法"就两种都写）
- 结果必须与**公认值**对比，给出**相对误差**和**误差来源分析**
- 算式写进答案正文，不要只丢一个最终数字

### 4. 画图

- **有完整数据表**（多组电压/电流等）→ 用 Node + sharp 走 SVG→PNG，画**一条连续 polyline**（断开会被用户立刻发现）
- **只有几个特征量**（如 X1 / X2 / ΔL）→ **不画图**，直接套公式计算。别为了走流程硬造一张图
- ❗图必须**嵌进最终文档**；只生成一张独立图片文件不算完成
- 插图尺寸：`LockAspectRatio=$false`，显式给 `Width`/`Height`（`Height = Width × 图高/图宽`）。
  **别用眼睛判断"图是不是太小"** —— 直接查：`pdfimages -list 报告.pdf`，看 `x-ppi`。
  宽 = 图宽像素 / x-ppi（英寸），与正文宽比一下就知道。

### 5. 填 Word（全程 COM）

可复用脚本见 `scripts/build-report.ps1`。铁律：

- 复制模板再改，**绝不直接改 `original.docx`**；改前先 `Set-ItemProperty -Name IsReadOnly -Value $false`
- 插入答案用 `$endBefore = $para.Range.End`（**不是** `$endBefore + 1`）
- 只填**数据处理 / 结论 / 讨论题**三节（预习页和数据页已经是照片了）

**章节编号用手写，不要跟 Word 自动编号死磕**：
删掉前面的段落之后，各标题的自动编号会各自从 1 开始（出现两个「一．」）；
`ApplyListTemplate(..., ContinuePreviousList=$true)` 也连不起来（实测三遍都没用）。三星期的功夫不如一行文本：

```powershell
foreach ($pair in @(@("数据处理","一．"), @("实验结论及现象分析","二．"), @("讨论题","三．"))) {
    $p = GetPara $doc $pair[0]
    try { $p.Range.ListFormat.RemoveNumbers() } catch {}
    $p.Range.InsertBefore($pair[1] + " ")     # 编号文本继承标题字体
    $p.Format.LeftIndent = 0; $p.Format.FirstLineIndent = 0
}
```

**公式上标**：正文里一律写 `10^-34`、`mv^2/2` 这种 `^` 写法（**不要用 Unicode ⁻¹⁹**，Times New Roman 缺字，PDF 里会渲染成 `10 * 19`），最后用 python-docx 转成真正的上标 run，并把里面的 `-` 换成 **U+2212 减号**（ASCII 连字符后允许断行，会把 `10⁻³⁴` 劈成两行）。

### 6. 一次成型清单 —— 开跑前就把这些写进脚本

下面每一条都是踩过的坑，**不要等第一版出来再逐页发现**：

| 项目 | 一次到位的写法 |
|---|---|
| 段落格式 | 插入段普遍会继承锚点编号 → 先 `ListFormat.RemoveNumbers()`；再设 `LineSpacingRule=1`(1.5 倍)、`SpaceBefore=6`、`CharacterUnitFirstLineIndent=2`(首行缩进 2 字符)、`Bold=$false`、`Size=Normal 字号` |
| 标题 | 每个章节标题设 `KeepWithNext=-1`，否则会孤零零掉在页底 |
| 签字框 / 浮动表格 | `Rows.WrapAroundText=$false` 取消环绕，否则会压在正文上 |
| 表标题间距 | 第二、三张表的标题段 `SpaceBefore=14` |
| 空段落 | 删之前查 `$p.Range.Information(12)`（表格内）+ `InlineShapes.Count`（含图片），否则会删坏表格或删掉照片 |
| 曲线图 | 段 `Alignment=1`(居中) + `KeepWithNext=-1`(与图注同页)；图注 9pt、居中 |
| 照片 | `LockAspectRatio=$true`、`Width=$availW`；第二张起 `PageBreakBefore=$true` |

### 7. 转 PDF

```powershell
$doc.SaveAs2($docx, 16)             # 16 = docx
$doc.ExportAsFixedFormat($pdf, 17)  # 17 = PDF，不改变文档身份
```

若用了 `^`→上标的 python 后处理：SaveAs2 → Close → 跑 python → 重开 → 导出 PDF（python-docx 保存过的文件 Word 会认）。

### 8. 验收

- [ ] PDF 页数 = 照片张数 + 正文页数；`InlineShapes.Count` = 照片数 + 曲线数
- [ ] **最前面每一页都是照片原图**（没有一页是打字重排）
- [ ] 每条答案紧跟对应题目下方、顺序正确
- [ ] 章节编号连续（一．二．三．），无重复编号
- [ ] 正文里看不到 `10^-34` 这种原始写法（说明上标转换漏了）
- [ ] 无 `【答】` 类机械标签、无免责话术
- [ ] 文件名 = 实验名称
- [ ] 中间文件已清理

## 高频坑

完整清单见 **`references/pitfalls.md`**（22 条，都是实战踩出来的）。最高频的几条：

| 症状 | 解法 |
|---|---|
| `Save()` 卡死、无报错 | 源 docx 带只读属性 → `Set-ItemProperty -Name IsReadOnly -Value $false` |
| 答案把题目劈成两半 | `Find` 返回的是匹配文本的 End → 用 `$f.Parent.Paragraphs.Item(1).Range.End` |
| 答案吞掉下一题题干 | 插入点偏移 1 字符 → 用 `$endBefore` |
| 照片钻进标题段落 | 图1 插好立刻 `InsertParagraphAfter()`；图2 先 `InsertParagraphBefore()` 造空段落当落点 |
| 文档多出一整页空白 | 手工分页符的锅 → 改用段落 `PageBreakBefore` |
| 章节编号乱窜 / 出现两个「一．」 | 删段后自动编号各自从 1 起 → 改**手写编号**（见流程 5） |
| `Word 未能引发事件` | 沙箱受限令牌 → 必须完全权限（ConstrainedLanguage 禁止创建 COM 对象）|

## 用户的硬性要求

完整版见 **`references/requirements.md`**。核心八条：

1. **最前面几页是手写照片原件**（一张一页，张数按用户给的），不是你重新打的字
2. 答案贴在**对应题目正下方**、按题号顺序 —— 不能集中到末尾
3. 不要 `【答】`、`答：` 这类机械标签
4. **严格沿用原文档格式**（字体、字号、标题样式、编号体系）
5. 只做被要求的事，不擅自加章节、不主动"优化"
6. 交付物**唯一**且命名正确
7. 报告里只写采用的那一套数据与结论；**疑问只在对话里说**
8. 素材（曲线图等）更新后**必须重跑填充脚本**，否则文档里还是旧图
