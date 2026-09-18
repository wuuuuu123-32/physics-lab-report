# physics-lab-report

给 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness)（dsh）用的一个 **Skill**：把大学物理实验的原始材料，做成一份可直接提交的报告 PDF。

## 它解决什么

做完一个物理实验，你手上通常是：几张手写照片、一张空白报告模板、一堆要算的东西。
这个 skill 把「**原始材料 → 可提交 PDF**」这条流水线固化下来——读手写数据、按公式计算、画图、填进原文档、检查格式、导出并按实验名称命名。

## 装法

**用户级安装**（任何工作区都能用）：

```powershell
git clone https://github.com/<你的用户名>/physics-lab-report.git "$env:USERPROFILE\.dsh\skills\physics-lab-report"
```

**项目级安装**（只对某个项目生效）：

```powershell
git clone <仓库地址> "<项目根>\.dsh\skills\physics-lab-report"
```

装完重启 `dsh`，它就会出现在 agent 的能力目录里，遇到匹配的任务自动加载。

## 结构

```
SKILL.md                      主指令（带 description 触发条件）
references/pitfalls.md        22 条实战易错点
references/requirements.md    交付要求（格式、命名、沟通偏好）
scripts/build-report.ps1      Word 填充脚本模板（改配置区即可用）
```

`SKILL.md` 保持精简以节省上下文，细节都放在 `references/` 里按需读取。

## 特色：踩坑驱动，不是规范推演

`SKILL.md` 里的每一条解法都对应一次**真实失败**，不是从文档里抄的：

| 症状 | 真正的原因 |
|---|---|
| `Save()` 静默卡死 300 秒 | 源 docx 带只读属性 |
| 答案把题目劈成两半 | `Find` 返回的是匹配文本的 End，不是段落 End |
| 照片钻进标题段落 | 靠字符偏移算插入点必错；应该用段前分页 + 造落点 |
| 章节编号乱窜成一…十、三/五/六/七 | Word 自动编号是段落属性，会被新段落继承 |
| Word COM 报 `Cannot create type` | 受限沙箱把 PowerShell 降到了 ConstrainedLanguage |
| 报告里出现「若采用另一个数值则…」 | AI 味，读者一眼看穿——问题应只在对话里说 |

## 三条设计原则

1. **能不写就不写** —— 一句话加不加，如果不影响要求、结果、可复现性，就删掉
2. **绝不编造数据** —— 手写看不清就重新切片放大；仍看不清就留空**并告诉用户**
3. **疑问只在对话里说** —— 报告正文只呈现采用的那一套数据与结论

## 适用 / 不适用

| ✅ 适用 | ❌ 不适用 |
|---|---|
| 手写照片 + 空白 docx → 填好的报告 PDF | 没有实验数据的普通写作 |
| 修订已完成报告里的数据或格式 | 纯排版任务 |
| 重复性的「大学物理实验报告」工作流 | 需要全新版式设计的文档 |

## 平台说明

面向 **Windows** 编写（Word COM 自动化、PowerShell 脚本）。数据读取、计算、画图部分与平台无关，但填 Word 那一步依赖本机装有 Office。

## License

MIT
