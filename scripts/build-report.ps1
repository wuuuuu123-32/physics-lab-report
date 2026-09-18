# ============================================================
#  大学物理实验报告 · Word 填充脚本模板
# ------------------------------------------------------------
#  用法：
#    1. 复制本文件到实验目录，改下面「配置区」
#    2. 执行（.ps1 不能直接跑）：
#         $code = Get-Content .\build-report.ps1 -Raw -Encoding UTF8
#         Invoke-Expression $code
#    3. 必须在「完全权限」沙箱下运行 —— Word COM 在受限模式下必然失败：
#         Cannot create type. Only core types are supported in this language mode.
#
#  目录约定：
#    <工作目录>\original.docx   原始空白模板（保持原样，脚本只复制不改它）
#    <工作目录>\page1.jpg       手写照片（预习页）      —— 没有就设 $null
#    <工作目录>\page2.jpg       手写照片（原始数据页）  —— 没有就设 $null
#    <工作目录>\curve-final.png 曲线图                  —— 没有就设 $null
# ============================================================

$ErrorActionPreference = "Stop"

# ══════════════════ 配置区：只改这里 ══════════════════

$work    = "D:\dsh\config\exp-report"        # 工作目录
$expName = "实验10 夫兰克-赫兹实验"           # 实验名称 → 决定输出文件名
$heading = "数据处理"                          # 正文起始标题（它之前的内容会被整页删除）
$titles  = @("数据处理", "实验结论及现象分析", "讨论题")   # 要保留自动编号的章节标题

# 照片（没有写 $null）
$photo1 = "$work\page1.jpg"
$photo2 = "$work\page2.jpg"

# 答案：题目定位片段 → 答案正文
#   定位片段必须在文档里唯一，且**避开编号和空格**（用题干里的独特中文片段最稳）
$answers = [ordered]@{
  "1. 题目关键字甲" = "答案正文……"
  "2. 题目关键字乙" = "答案正文……"
}

# 插图：题目定位片段 → 图片路径
$figures = [ordered]@{
  "利用计算机软件绘制" = "$work\curve-final.png"
}

# ══════════════════ 配置区结束 ══════════════════

$out = "$work\$expName.docx"
$pdf = "$work\$expName.pdf"
$src = "$work\_src.docx"

# ---------- 工具函数 ----------

function GetPara($doc, $text) {
    $f = $doc.Content.Find
    $f.ClearFormatting(); $f.Text = $text; $f.Forward = $true; $f.Wrap = 0
    if ($f.Execute()) { return $f.Parent.Paragraphs.Item(1) } else { return $null }
}

function ParaStart($doc, $text) {
    $p = GetPara $doc $text
    if ($p) { return $p.Range.Start } else { return $null }
}

# 在指定段落之后插入一个新段落（❗必须用 $endBefore，不是 $endBefore + 1）
function AppendParaAfter($doc, $para, $text) {
    $endBefore = $para.Range.End
    $para.Range.InsertParagraphAfter()
    $doc.Range($endBefore, $endBefore).InsertAfter($text)
}

function AppendPicAfter($doc, $para, $imgPath, $maxWidth) {
    $endBefore = $para.Range.End
    $para.Range.InsertParagraphAfter()
    $pic = $doc.InlineShapes.AddPicture($imgPath, $false, $true, $doc.Range($endBefore, $endBefore))
    $pic.LockAspectRatio = $true
    if ($pic.Width -gt $maxWidth) { $pic.Width = $maxWidth }
    $doc.Range($pic.Range.End, $pic.Range.End).InsertParagraphAfter()
    return $pic
}

# ---------- 主流程 ----------

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2
foreach ($f in @($out, $pdf, $src)) { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force } }

Copy-Item "$work\original.docx" $src -Force
# ❗不解除只读会导致 Save() 静默卡死 300 秒
Set-ItemProperty -Path $src -Name IsReadOnly -Value $false

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($src)

    $availW = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin

    # ── ① 先填文档后半部分（位置在后面，不受前面删除影响）──
    foreach ($key in $answers.Keys) {
        $p = GetPara $doc $key
        if ($p) { AppendParaAfter $doc $p $answers[$key] }
        else { Write-Host "!! 未找到题目：$key" }
    }
    foreach ($key in $figures.Keys) {
        $p = GetPara $doc $key
        if ($p -and (Test-Path $figures[$key])) { AppendPicAfter $doc $p $figures[$key] $availW | Out-Null }
        else { Write-Host "!! 未找到题目或图片：$key" }
    }
    Write-Host "答案与插图已填 ✓"

    # ── ② 删掉正文起始标题之前的全部内容（整页物理删除）──
    $sProc = ParaStart $doc $heading
    if ($null -ne $sProc -and $sProc -gt 0) {
        $doc.Range(0, $sProc).Delete() | Out-Null
        Write-Host "已删除「$heading」之前的全部内容"
    } else { Write-Host "!! 未找到起始标题：$heading" }

    # ── ③ 照片放最前面，每张各占一页 ──
    if ($photo1 -and (Test-Path $photo1)) {
        $pic1 = $doc.InlineShapes.AddPicture($photo1, $false, $true, $doc.Range(0, 0))
        $pic1.LockAspectRatio = $true; $pic1.Width = $availW
        # 立刻分段，否则原标题会被挤进图片所在的段落
        $e1 = $pic1.Range.End
        $doc.Range($e1, $e1).InsertParagraphAfter()

        if ($photo2 -and (Test-Path $photo2)) {
            $tPara = GetPara $doc $heading
            if ($tPara) {
                $tPara.Range.InsertParagraphBefore()      # 造一个空段落当落点
                $holder = $tPara.Previous()
                $pic2 = $doc.InlineShapes.AddPicture($photo2, $false, $true, $holder.Range)
                $pic2.LockAspectRatio = $true; $pic2.Width = $availW
                $holder.PageBreakBefore = $true           # 图2 另起一页
                $tPara.PageBreakBefore = $true            # 正文另起一页
            }
        }
        Write-Host "照片已置于文档最前"
    }

    # ── ④ 只保留章节标题的自动编号（否则编号会乱窜）──
    $cleared = 0
    foreach ($p in $doc.Paragraphs) {
        $t = ($p.Range.Text -replace "[\r\n\a\x07]", "").Trim()
        $isTitle = $false
        foreach ($k in $titles) {
            if ($t -like "*$k*" -and $t.Length -le ($k.Length + 6)) { $isTitle = $true; break }
        }
        if (-not $isTitle) { try { $p.Range.ListFormat.RemoveNumbers(); $cleared++ } catch {} }
    }
    Write-Host "已清除 $cleared 个段落的编号"

    # ── ⑤ 另存 docx + 导出 PDF ──
    $doc.SaveAs2($out, 16)                        # 16 = wdFormatDocumentDefault
    Write-Host ("页数 " + $doc.ComputeStatistics(2) + "  图片 " + $doc.InlineShapes.Count)
    $doc.ExportAsFixedFormat($pdf, 17)            # 17 = wdExportFormatPDF
    Write-Host ("PDF: $pdf  " + (Get-Item -LiteralPath $pdf).Length + " 字节")

    # ── ⑥ 打印最终结构，便于核对 ──
    Write-Host "=== 最终结构 ==="
    foreach ($p in $doc.Paragraphs) {
        $t = ($p.Range.Text -replace "[\r\n\a\x07]", "").Trim()
        $pic = if ($p.Range.InlineShapes.Count -gt 0) { "[图]" } else { "" }
        if ($t.Length -gt 0 -or $pic -ne "") {
            $num = $p.Range.ListFormat.ListString
            $pg = [string]$p.Range.Information(3)
            if ($t.Length -gt 38) { $t = $t.Substring(0, 38) + "…" }
            Write-Host ("  P" + $pg + " " + $num + " " + $pic + " " + $t)
        }
    }

    $doc.Close(0)
    $word.Quit()
    Remove-Item -LiteralPath $src -Force -ErrorAction SilentlyContinue
    Write-Host "完成 ✓"
} catch {
    Write-Host ("失败: " + $_.Exception.Message)
    try { $doc.Close(0) } catch {}
    try { $word.Quit() } catch {}
}
