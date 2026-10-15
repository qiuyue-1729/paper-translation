# Paper Translation

将各学科学术论文完整翻译为中文，提炼关键内容，并使用用户模板或内置默认模板生成 Word。

## 安装

将仓库中的 `paper-translation` 目录复制到 `~/.codex/skills/`，重新加载后使用 `$paper-translation` 调用。

## 使用

提供论文文件并说明翻译范围。默认每篇生成一份中文 Word，关键内容提炼在前、全文译文在后，保留公式、原图和参考文献。优先遵循当次用户要求及模板，没有时使用内置模板。正文采用正文样式，独立公式采用公式样式。

## 公式插件

已提供 `paper-translation/scripts/convert-equations.ps1`，通过 Word COM 调用本机 AxMath 或 MathType 的 Word 宏，将独立段落中的 TeX 转成可编辑公式。需要 Windows、Word 和可用的插件安装；详细用法及边界见 `references/equation-plugins.md`。

## 验证范围

默认模板已用 Word 导出并检查空白页及两页填充样张。AxMath 2.5 与 MathType 7.4.2 已完成四类公式的转换、保存重开和反向 TeX 读取；最终脚本使用模板公式样式，正文样式保留。结果适用于所测环境，每篇正式译文仍需核对公式和逐页检查。
