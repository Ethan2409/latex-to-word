# latex-to-word

LaTeX 公式转 Word 论文格式的 VBA 宏，自动处理斜体/正体/上下标。

## 解决什么问题

从 AI 页面（ChatGPT、Copilot 等）复制 LaTeX 公式到 Word，结果要么是原始的 `$...$` 字符串，要么格式全乱——该斜体的没斜体、该正体的没正体、上下标塌成一行。

这两个宏把 `$...$` 包起来的 LaTeX 内联公式，自动转成 Word 论文要求的格式：

- 变量（字母、希腊字母）→ 斜体
- 标准数学函数（sin / cos / log / ln / exp / lim / max / min）→ 正体
- `_` 下标、`^` 上标 → 真正的上下标
- `\times`、`\alpha` 等命令 → Unicode 符号

比如 `$R_i$` 转成 R（大写斜体）+ i（下标正体）。

## 两个宏

| 宏名 | 用法 |
|---|---|
| [PasteAsTextAndConvertLatex](macro/PasteAsTextAndConvertLatex.bas) | 先从把ai生成的内容复制一下（按crtl+C），再运行宏（自动粘贴并转换） |
| [ConvertLatexInSelection](macro/ConvertLatexInSelection.bas) | 先从把ai生成的内容以纯文本形式复制到word，然后选中要转换的段落，再运行宏 |

两个宏转换逻辑完全相同，只是输入源不同：一个吃剪贴板，一个吃选区。

## 安装

1. 把 [Latex约束](macro/Latex约束.md) 添加到ai智能体(如codex或claude)项目下的AGENTS.md中
2. 打开 Word，`Alt` + `F11` 打开 VBA 编辑器
3. 插入 → 模块
4. 把 `macro/` 下对应 `.bas` 文件的内容粘贴进去
5. 给宏设置快捷键（方法很简单，可以直接百度）
6. 回到 Word 使用

## 转换规则（支持的范围）

- 内联公式：`$...$`（单美金）
- 符号：`-` → 标准减号、`\times` → ×，以及希腊字母 `\Phi` `\phi` `\alpha` `\beta` `\gamma` `\theta` `\Delta` `\delta` `\mu` `\rho` `\sigma` `\eta` `\varepsilon`
- `\text{...}` / `\mathrm{...}` 自动剥离，只留内容
- 字母 + 希腊字母自动斜体；数学函数自动正体
- 下标 `_{...}` 或 `_x`；上标 `^{...}` 或 `^x`

## 已知限制

- `$$...$$`（双美金）独立公式块会被跳过，不转换
- 只覆盖上面列出的符号和希腊字母，`\frac`、`\sqrt` 等其他 LaTeX 命令不在处理范围内

## License

MIT

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=Ethan2409/latex-to-word&type=Date)](https://star-history.com/#Ethan2409/latex-to-word&Date)
