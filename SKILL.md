---
name: phd-thesis-formatter
description: 当用户需要把博士学位论文 .docx 按学校模板做格式检查、版式规范化或送审前排版整理时使用，尤其是提到“论文格式””博士论文格式化””学位论文模板””图书馆版””匿名送审版””三线表””页眉页脚””目录””Nature 参考文献””标点混用””中英文空格”或”词语重复”时。
---

# 博士学位论文格式化

本 skill 适用于接近定稿的 `.docx` 学位论文。

它是一个“半自动格式化 + 明确人工复核项”的工作流，不应被表述为“任何论文都能一键完全合规”。

详细格式参数见 `references/format_specifications.md`。
高风险人工复核项见 `references/manual_review_checklist.md`。

## 能力边界

### 可以安全自动处理的项目

- 页面大小、页边距、页眉页脚距离等节属性
- 已有正文 run 的中英文字体归一化
- 普通表格转换为三线表边框
- 基础结构检查：页面参数、字体统计、图表/参考文献是否被识别到

### 不能可靠全自动完成的项目

- 无编号标题的层级判断
- 复杂分节、奇偶页不同页眉、`STYLEREF` 章节名联动
- 目录域更新
- 匿名送审版脱敏
- 从混杂格式、APA、BibTeX 或不完整文本准确转换为 Nature 参考文献
- 正文中引用序号和参考文献列表的重新映射
- 跨页续表、分图、复杂公式、上下标和特殊符号的逐项校正

## 输入要求

- 输入文件必须是 `.docx`
- 默认处理对象是“内容已基本完成”的博士论文或接近终稿的章节
- 开始前必须确认：
  - 是图书馆最终版，还是匿名送审版
  - 这次目标是“基础格式规范化”，还是“逐项冲最终合规”
  - 参考文献是否已有可解析的结构化来源

## 配置

首次使用前，编辑 `scripts/thesis_config.py` 填入你的学校信息：

```python
SCHOOL_NAME = 'YOUR_UNIVERSITY'          # 例如 '北京大学'
EVEN_PAGE_HEADER = SCHOOL_NAME + '博士学位论文'
MIN_PAGES_BEFORE_REFERENCES = 100
REQUIRED_POST_REF_SECTIONS = ['致谢', '原创性声明', '使用授权说明']
```

## 工作流程

### 1. 先做结构审查

- 识别是否包含：封面、版权声明、中英文摘要、目录、正文、参考文献、附录、致谢、原创性声明
- 判断是否存在明显的分节边界：摘要、目录、正文、附录、致谢
- 如果用户要做匿名送审版，先把脱敏要求单独列出来，不要和普通图书馆版混为一谈

### 2. 解包并备份

如果环境里已安装 `docx` skill，优先复用其 Office 脚本：

```bash
DOCX_SKILL_DIR="${CODEX_HOME:-$HOME/.codex}/skills/docx"
python "$DOCX_SKILL_DIR/scripts/office/unpack.py" input.docx unpacked/
```

如果只是执行低风险自动规范化，优先直接运行仓库内主脚本：

```bash
python scripts/auto_format_docx.py input.docx --output output.docx --report format_report.txt
```

### 3. 只自动修正“低风险项”

对 `unpacked/word/document.xml` 或对应 XML 做以下安全修改：

- 页面尺寸设为 A4
- 页边距设为：上 3.0 cm、下 2.5 cm、左/右 2.6 cm
- 页眉距边界 2.0 cm，页脚距边界 1.75 cm
- 普通表格转三线表
- 对现有 run 缺失的中英文字体做基础补全
- 对证据明确的编号标题补上 `Heading1` / `Heading2` / `Heading3` 样式标记

表格转换脚本：

```bash
python scripts/convert_to_three_line_table.py unpacked/word/document.xml
```

### 4. 谨慎处理高风险项

只有在证据充分时才自动修改：

- 只有明确出现“第一章 / 第二章 ...”时，才能判定为章标题
- 只有明确出现 `1.1`、`1.1.1` 这类编号时，才能提升为节标题
- 如果参考文献源数据不完整，不要声称“已转成 Nature 且完全准确”
- 如果没有完整分节逻辑，不要声称“页眉页脚已完全符合奇偶页规则”

### 5. 行内批注审核（可选）

生成带行内红色批注的审核版本，检查包括：中英文标点混用、中英文之间缺少半角空格、词语重复、数字与单位之间缺少空格、图表未在正文中引用等：

```bash
python scripts/generate_school_audit_inline_notes.py input.docx --output input_audit.docx
```

### 6. 生成行内审核版 + 保守自动修复版（可选）

同时输出两个文件和一份报告：

```bash
python scripts/generate_review_variants.py input.docx \
  --inline-output input_review.docx \
  --fixed-output input_fixed.docx \
  --report review_report.txt
```

### 7. 重新打包并验证

```bash
python "$DOCX_SKILL_DIR/scripts/office/pack.py" unpacked/ output.docx --original input.docx
python scripts/validate_format.py output.docx
```

## 输出要求

至少输出两个结果：

1. `[原文件名]_formatted.docx`
2. 一份格式化报告，必须拆成三类：
   - 已自动修复
   - 已检测但需要人工复核
   - 未处理/不能安全自动处理

## 结果表述规则

- 只有在验证结果通过，并且 `references/manual_review_checklist.md` 中的高风险项已逐项核对后，才可以说“基本符合要求”或“可提交终稿”
- 不要把“页面设置、基础字体、三线表已处理”表述成“全部格式已完全符合要求”
- 对参考文献、目录、匿名版脱敏、复杂页眉页脚，默认使用“需要人工复核”措辞

## 推荐使用顺序

1. 读取 `references/format_specifications.md`
2. 解包 `.docx`
3. 先修复页面与表格等低风险项
4. 再决定是否处理标题、页眉页脚、参考文献
5. 运行 `scripts/validate_format.py`
6. 对照 `references/manual_review_checklist.md` 出报告
