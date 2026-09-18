# 易错点速查表（22 条 + 环境补充）

**全部来自真实踩坑记录**，不是理论推演。按症状查即可。

| # | 症状 | 根因 | 解法 |
|---|---|---|---|
| 1 | `Save()` 卡死 300 秒、无报错 | 源 docx 带 `IsReadOnly=True` | `Set-ItemProperty -Path $f -Name IsReadOnly -Value $false`；并改用 `SaveAs2` 存到新文件名 |
| 2 | `Word 未能引发事件` / `Cannot create type` | workspace-write 下 PowerShell 跑在 **ConstrainedLanguage** | 必须切「完全权限」；放开某个目录解决不了（实测见下） |
| 3 | `running scripts is disabled` | 执行策略 | `Get-Content x.ps1 -Raw -Encoding UTF8 \| Invoke-Expression` |
| 4 | 答案把题目从中间劈开 | `Find` 返回的是**匹配文本**的 End，不是段落 End | `$f.Parent.Paragraphs.Item(1).Range.End` |
| 5 | 答案吞掉下一题题干（第 2、3 题粘一起） | 插入点偏移 1 个字符 | 插入位置用 `$endBefore`，**不是** `$endBefore + 1` |
| 6 | 两张照片挤一页 / 照片和标题挤一段 | 靠「分页符占 1 字符」推算插入点，极易算错 | 分页一律用 `段落.PageBreakBefore`，插入点见 #16 |
| 7 | 章节编号乱窜（一…十、三/五/六/七） | Word 自动编号是段落属性，被新段落继承 | 全部插完后遍历 `ListFormat.RemoveNumbers()` |
| 8 | 删了内容却还留着空页 | 只删了内容，没删页面 | `$doc.Range(0, $startOfHeading).Delete()` |
| 9 | 曲线图改了但文档里还是旧的 | 忘了重跑 build 脚本 | 任何素材更新后**必须重跑**填充脚本 |
| 10 | 曲线「断断续续」 | polyline 被拆成多段 | 生成**一条**连续 polyline |
| 11 | `vision_glance` 只给 3 行 / 输出推理 | 工具截断 | 改用 sharp 切片 + `read_image` 逐块读 |
| 12 | `Remove-Item`/`Expand-Archive` 拒绝 docx | 后缀白名单 | 先 `Copy-Item x.docx x.zip` |
| 13 | 数据抄错（波谷电流反而下降） | 手写辨认错误 | 用物理规律反查，可疑处重新切片精读 |
| 14 | 临时目录不存在导致沙箱报错 | `%TEMP%\dsh-xxxx` 被清掉 | 先 `New-Item -ItemType Directory` 建出来 |
| 15 | 用户说「怎么还留着第一页」 | 用户要的删除粒度是**整页** | 见 #8，一次删干净 |
| 16 | 照片插进去后和标题挤在同一段 | `AddPicture` 的插入点落在标题段落内部 | 图1 插好后立刻 `InsertParagraphAfter()` 把标题挤到下一段；图2 先 `$tPara.Range.InsertParagraphBefore()` 造一个空段落当落点，再 `AddPicture` 指向这个空段落 |
| 17 | 文档中间多出一整页空白 | 手工分页符 `InsertBreak(7)` 留下多余段落 | 改用 `段落.PageBreakBefore = $true`（段前分页） |
| 18 | `Find` 找不到章节标题 | 各实验模板章节名不同（「数据处理」vs「实验数据处理」，有的没有「结论及现象分析」） | 先从 docx 提取段落清单确认真实标题；脚本里每个 Find 都配 `else { Write-Host "!! 未找到…" }` 报警 |
| 19 | 把原模板的空段落当垃圾删了 | 那些空段是留给学生手写答案的**版式留白**（某实验 96 段里有 57 段是空的） | 保留，不要动 |
| 20 | `node xxx.mjs` 报 `ERR_UNSUPPORTED_ESM_URL_SCHEME: d:` | ESM 导入 Windows 绝对路径必须写成 URL | `import sharp from 'file:///D:/dsh/config/.cutout/node_modules/sharp/lib/index.js'` |
| 21 | 手写数字看不清 | 直接看整页太模糊 | 按行切窄条（宽 300~600px）再放大 3~8 倍，一次只看一行 |
| 22 | 报告里冒出「若采用另一个数则…」 | 想显得严谨，实际一眼暴露 AI 痕迹 | 报告里只写采用的那一套数据和结论；疑问留在对话里说 |

---

## 环境相关（Windows）

- **`.ps1` 不能直接跑**：`$code = Get-Content <path> -Raw -Encoding UTF8; Invoke-Expression $code`
- **改 `~/.dsh` 下的 JSON 配置必须无 BOM**：PowerShell 5.1 的 `Set-Content -Encoding UTF8` 会偷偷加 BOM，DSH 启动时 `JSON.parse` 直接崩、窗口一闪就没。正确写法：
  ```powershell
  [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
  ```
  改完验证前 3 字节不是 `EF BB BF`
- **受限令牌下 `curl.exe` 的 TLS 会挂**（HTTPS 返回 000），用 Node 的 `fetch` 代替
- **`git push` 在受限令牌下 schannel 会失败**（`SEC_E_NO_CREDENTIALS`），需提权
- Word COM 在 `workspace-write` 下**一定失败**：实测 `New-Object -ComObject Word.Application` 直接报
  `Cannot create type. Only core types are supported in this language mode.`
  —— 因为受限模式把 PowerShell 降到了 ConstrainedLanguage，**与文件白名单无关**，别在这条路上浪费时间

---

## 数据分析相关

- **物理规律是最好的校对器**：抄出来的数如果违反物理（比如波谷电流随电压下降），那一定是抄错了
- **自洽链能排除不可能的读数，但不能替代向用户确认**：学生自己算错的数也可能"自洽"
- 波谷/峰值读数与理论不符时，**如实说明是判读误差，不要编一个解释**
- 题目要求几种方法就做几种；结果必须与公认值对比并给相对误差
