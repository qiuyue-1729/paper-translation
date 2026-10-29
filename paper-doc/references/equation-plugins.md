# AxMath 与 MathType 调用

## 已实现的入口

（1）使用 [转换脚本](../scripts/convert-equations.ps1) 通过 Word COM 调用本机插件。AxMath 使用 AMCCTex2AMEqn.ConvertTeX2AMEqn；MathType 使用 MTCommand_TeXToggle。入口从已安装插件随附的 Word 宏模板确认，脚本不包含厂商代码或二进制。

（2）需要 Windows、Word、可正常使用的 AxMath 或 MathType，以及 PowerShell 7。插件宏需已加载；可通过 AddinPath 指定本机安装附带的可信 Word 加载项。不得为运行而降低宏安全设置。优先 AxMath，不可用时切换 MathType 或 Word 原生公式。

## 调用方式

```powershell
pwsh -File scripts/convert-equations.ps1 -InputDocx staging.docx -OutputDocx converted.docx -Engine AxMath -ExpectedCount 4 -OutputPdf converted.pdf
pwsh -File scripts/convert-equations.ps1 -InputDocx staging.docx -OutputDocx converted.docx -Engine MathType -ExpectedCount 4 -OutputPdf converted.pdf
```

（1）输入使用包含模板样式的独立中间文档，公式各占一个段落，采用 $...$、$$...$$、\(...\) 或 \[...\] 定界。脚本从后向前逐公式转换，避免整篇混合字号影响插件字号。不要将已有公式对象的成品传入：MathType Toggle 为双向转换，重复调用可能把对象变回文本。正文行内简单变量直接使用 Word 上下标。

（2）正文使用模板的“正文”样式；独立公式使用已有“公式”样式，可用 EquationStyle 参数指定用户模板中的对应名称。转换后重新应用该样式并清除插件增加的段落直接格式。保留正文样式及行距，不再逐公式写死行距。默认字号继承模板；仅在用户明确要求时传入 FontSize。

（3）输入和输出路径必须不同，输出必须尚不存在。脚本核对转换前后数量、OLE 类型和保存重开后的对象数量；失败输出只是诊断副本，不能作为交付。脚本仅覆盖主文档中的独立公式段落，不处理混合行内复杂公式、文本框、页眉页脚或已存在的图形对象。公式编号在转换后按模板添加，避免混入转换段落。

（4）插件出现许可、宏被阻止或交互对话框时保留错误，解决实际环境问题后再处理；不要盲目重复调用。新模板缺少公式样式时根据用户要求补齐或映射，不能自动覆盖正文样式。

## 已取得的运行结果

2026-10-15 在 Word 与 AxMath 2.5、MathType 7.4.2 环境中，下标、分式、矩阵、多层上下标四类样例均已实际转换。生成的类型分别为 Equation.AxMath 和 Equation.DSMT4，保存重开后四个对象保持存在。另对保存对象执行反向 TeX 转换，四个对象均还原为有内容的公式，证明不是图片包装。已检查生成页面，并修正了混合字号与正文固定行距造成的问题；最终脚本改用模板“公式”样式。

这些结果证明所测版本和样例的调用可用，不保证任意 LaTeX 命令、其他版本或机器环境均兼容。每篇正式输出仍应核对公式语义和最终页面。

## 安装来源

优先使用本机安装，不为翻译任务自动升级。需要安装时使用官方入口或用户提供的可信安装程序，并按当次授权处理。AxMath：https://www.axsoft.co/ 。MathType：https://www.wiris.com/en/mathtype/ 。本 skill 不附带商业安装包或激活材料。
