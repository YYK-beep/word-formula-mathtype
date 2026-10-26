# word-formula-mathtype

将 Word 写作中常见的公式来源统一转换为真正的 **MathType OLE 对象**
（`Equation.DSMT4`）。该 skill 面向 Windows 桌面版 Word，强调真实
MathType 对象、转换前自检、可复核 OCR、批量处理和结果验证。

> 这里说的“任意形式”指本 skill 支持的输入形态：Word OMML、DOCX 内 LaTeX
> 文本、TeX/Markdown/HTML/MathML 文件、公式截图和公式照片。整页 PDF 或旧版
> Equation Editor 3.0 仍需先裁剪或提供人工审核后的 LaTeX。

## 适用场景

- 数学建模竞赛论文与支撑材料
- 本科/硕士/博士毕业论文、开题报告
- 综述论文、期刊论文、会议论文
- 课程设计、实验报告、项目申报书
- 需要将 Word、TeX、公式截图或照片统一整理为 MathType 的编辑场景

## 支持的公式来源

- Word 原生 OMML（`m:oMath`）
- DOCX 中的 LaTeX 文本：`$...$`、`$$...$$`、`\(...\)`、`\[...\]`
- 无定界符但可高置信识别的独立 LaTeX 段落
- `.tex`、Markdown、HTML、MathML 文件
- DOCX 内嵌的公式截图或照片
- 独立 PNG/JPG/BMP/TIFF/WebP 公式图片

## 开始前强制自检

统一转换脚本和旧版 OMML 脚本都会在开始处理前自动运行环境预检。预检失败会
返回非零退出码并停止后续操作，不会在缺少 Word 或 MathType 时悄悄降级。

可单独运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

机器可读结果：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1 -Json
```

预检内容：

- Windows PowerShell 5.1+
- Microsoft Word 桌面版 2016+
- MathType 7.x 注册信息与版本
- Word STARTUP 目录中的 `MathType Commands 2016.dotm`
- Python 3.10+
- Pandoc 3.x 或 `pypandoc-binary`
- 图片模式下所需的 VLM/Mathpix 配置，或 Manual OCR 映射

`-SkipPreflight` 只应在已经手动完成环境检查时使用。

## 支持条件

| 组件 | 最低要求 | 本次实测版本 | 说明 |
| --- | --- | --- | --- |
| Windows | Windows 10/11 | Windows 11 10.0.22631 | 仅支持 Windows |
| PowerShell | 5.1 | 5.1.22621.6133 | 脚本兼容 Windows PowerShell |
| Microsoft Word | Word 2016+ 桌面版 | Office 2021 16.0.20527.20088 | 不支持仅网页版 Word |
| MathType | MathType 7.x | MathType 7.12.3 | 必须有 Word 加载项 `MathType Commands 2016.dotm` |
| Python | 3.10+ | 3.12.3 | 用于 OMML 提取和 OCR 调用 |
| Pandoc | 3.x | 3.9 | 推荐安装 `pypandoc-binary` |
| 图片 OCR | 可选 | SenseNova `sensenova-6.8-flash-lite` | 也可使用 Mathpix 或 Manual OCR JSON |

图片没有 OCR 凭据时，仍可通过 `-OcrBackend Manual` 和人工审核的 LaTeX 映射
完成转换。

## 测试效果

完整测试材料位于 [`examples/formula-speed-test`](examples/formula-speed-test)：

- [测试源文本](examples/formula-speed-test/source.md)
- [测试输入 Word](examples/formula-speed-test/input.docx)
- [MathType 输出 Word](examples/formula-speed-test/output-mathtype.docx)
- [MathType 输出 PDF](examples/formula-speed-test/output-mathtype.pdf)
- [转换日志](examples/formula-speed-test/conversion-log.json)

测试文档包含 12 个公式：

- 6 个行内公式、6 个独立公式
- 6 个原有 OMML、6 个 Word 内 LaTeX 文本
- 覆盖分式、根式、矩阵、积分、求和、极限、分段函数、偏微分和希腊字母

实测结果：

| 指标 | 结果 |
| --- | ---: |
| 转换成功 | 12/12 |
| 最终 MathType 对象 | 12 |
| 剩余 OMML | 0 |
| 剩余公式定界符 | 0 |
| 剩余原始 LaTeX 段落 | 0 |
| 端到端总耗时（含强制预检） | 23.881 秒 |
| 单条公式平均转换耗时 | 0.700 秒 |
| 脚本内部固定开销（Word/Pandoc/收尾） | 约 12.9 秒 |

测试材料包括输入 Word、测试源文本、输出 MathType Word、输出 PDF、转换日志、
速度报告和页面预览图。

## 快速使用

### 自动处理 DOCX 中的 OMML、LaTeX 和公式图片

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -Mode Auto `
  -ImagePolicy Auto
```

### 只转换 Word 内 LaTeX 文本

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -Mode Latex `
  -ImagePolicy None
```

### 转换 TeX 文件

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\paper.tex" `
  -TargetPath "D:\docs\paper_mathtype.docx"
```

### 转换公式照片

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\formula.jpg" `
  -TargetPath "D:\docs\formula_mathtype.docx" `
  -Mode Images `
  -OcrBackend Vlm
```

### 先审核图片 OCR，再转换

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -Mode Images `
  -ImagePolicy Auto `
  -DryRun
```

修正或确认生成的 JSON 后：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -Mode Images `
  -OcrBackend Manual `
  -ManualOcrMapPath "D:\docs\reviewed_ocr.json"
```

## OCR 配置

默认 OpenAI 兼容 VLM：

```powershell
$env:FORMULA_OCR_API_KEY = "..."
$env:FORMULA_OCR_BASE_URL = "https://token.sensenova.cn/v1"
$env:FORMULA_OCR_MODEL = "sensenova-6.8-flash-lite"
```

也支持 SenseNova 环境别名：

- `SN_VISION_API_KEY` / `SN_CHAT_API_KEY` / `SN_API_KEY`
- `SN_VISION_BASE_URL` / `SN_CHAT_BASE_URL` / `SN_BASE_URL`
- `SN_VISION_MODEL` / `SN_CHAT_MODEL`

Mathpix：

```powershell
$env:MATHPIX_APP_ID = "..."
$env:MATHPIX_APP_KEY = "..."
```

## 验证输出

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify_mathtype.ps1 `
  -Path "D:\docs\input_mathtype.docx"
```

完成后的关键指标：

- `omathsInRange = 0`
- `mathTypeShapes > 0`
- `remainingFormulaDelimiters = 0`
- `remainingRawLatexParagraphs = 0`

`remainingDollarCharacters` 可能是正文中的货币符号，例如 `$5`，因此验证脚本会
将普通美元符号和公式定界符分开统计。

## English overview

`word-formula-mathtype` converts supported formula sources into genuine
MathType OLE objects in Windows desktop Word. It supports OMML, LaTeX text in
DOCX, TeX/Markdown/HTML/MathML files, screenshots, and formula photos.

A mandatory preflight runs automatically before conversion and checks Word,
MathType 7 registration/version, the MathType Word add-in, Python, Pandoc,
PowerShell, and optional OCR configuration. Missing required components stop the
job with a non-zero exit code.

See [`examples/formula-speed-test`](examples/formula-speed-test) for reproducible
input/output Word files, PDF output, logs, and benchmark results.

## Limitations

- 必须有 Windows 桌面版 Word 和 MathType 7；不支持仅网页版 Word。
- PDF 页面和旧版 Equation Editor 3.0 不会自动拆分。
- 手写或模糊公式建议先用 `-DryRun` 审核识别结果。
- 含中文 `\text{...}` 的复杂公式仍需人工复核。
- 不保证所有第三方 LaTeX 宏都受 MathType TeX 导入器支持。

## License

MIT
