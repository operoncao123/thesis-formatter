# thesis-formatter

博士学位论文 `.docx` 格式化与审核工具，同时支持 **Claude Code** 和 **OpenAI Codex**。

半自动工作流：低风险项自动修复，高风险项生成行内批注由人工复核。不承诺一键全自动合规。

---

## 功能特点

### 自动格式化

- A4 纸张、页边距、页眉页脚距离归一化
- 中英文字体基础补全（宋体 / Times New Roman / 黑体）
- 普通表格转三线表（去竖线、去内横线、加粗顶底线）
- 编号明确的标题自动补 Word 标题样式（Heading1 / Heading2 / Heading3）
- 输出格式化报告，分为：已自动修复 / 需人工复核 / 未处理

### 行内批注审核

在原文档对应段落下方插入红色批注，覆盖以下检查项：

| 检查项 | 说明 |
|--------|------|
| **中英文标点混用** | 检测中文语境中出现英文逗号、句号、分号等，提示统一使用中文标点 |
| **缺少半角空格** | 检测中文与英文字母/数字直接相邻（如 `NF-κB激活` → `NF-κB 激活`），自动跳过文献引用编号 `[1]` |
| **词语重复** | 检测相邻重复的中文词组（2–6 字）或英文单词，标注疑似笔误 |
| **数字与单位间距** | 检测 `10μm` 应写为 `10 μm` 的情况 |
| **英文直引号** | 中文行文中出现英文直引号时提示改为中文引号 |
| **无编号标题样式** | 使用了 Word 标题样式但没有编号，会导致目录链条断裂 |
| **图表编号重复** | 图号/表号与前文重复时提示 |
| **图表未在正文引用** | 检测图/表标题对应的正文引用是否存在 |
| **参考文献页数** | 检查参考文献前页数是否达到学校要求 |
| **缺少必要章节** | 检测致谢、原创性声明等是否存在 |

### 审核变体生成

一次命令生成两个版本：

- **行内批注版**：所有问题以红色注释插入段落下方，便于在 Word/WPS 中逐条确认
- **保守自动修复版**：只应用低风险修复（字体归一化、标题空格等）

### 需要人工复核的项目

以下内容工具不会自动修改，会在报告或批注中标出：

- 奇偶页页眉与多分节页码切换
- 目录域刷新
- 匿名送审版脱敏
- 参考文献从 APA/混杂格式准确转换为 Nature 格式
- 正文引用序号与参考文献列表的重新映射
- 跨页续表、分图、复杂公式与上下标

---

## 安装与使用

### 1. 克隆仓库

```bash
git clone git@github.com:operoncao123/thesis-formatter.git
cd thesis-formatter
pip install -r requirements.txt
```

### 2. 配置学校信息

编辑 `scripts/thesis_config.py`：

```python
SCHOOL_NAME = 'YOUR_UNIVERSITY'          # 例如 '北京大学'
EVEN_PAGE_HEADER = SCHOOL_NAME + '博士学位论文'
MIN_PAGES_BEFORE_REFERENCES = 100
REQUIRED_POST_REF_SECTIONS = ['致谢', '原创性声明', '使用授权说明']
```

### 3. 在 Claude Code 中使用

```bash
mkdir -p ~/.claude/skills
ln -s "$(pwd)" ~/.claude/skills/phd-thesis-formatter
```

之后在 Claude Code 中直接描述需求即可，例如：

> 帮我格式化这篇博士论文 thesis.docx，检查中英文标点、空格和词语重复问题

### 4. 在 OpenAI Codex 中使用

```bash
mkdir -p ~/.codex/skills
ln -s "$(pwd)" ~/.codex/skills/phd-thesis-formatter
```

`agents/openai.yaml` 已配置好 Codex 的接入描述。

### 5. 直接命令行运行

```bash
# 自动格式化
python scripts/auto_format_docx.py thesis.docx --output thesis_formatted.docx --report report.txt

# 行内批注审核（重点推荐）
python scripts/generate_school_audit_inline_notes.py thesis.docx --output thesis_audit.docx

# 行内批注版 + 保守自动修复版
python scripts/generate_review_variants.py thesis.docx \
  --inline-output thesis_review.docx \
  --fixed-output thesis_fixed.docx \
  --report review_report.txt

# 仅验证格式
python scripts/validate_format.py thesis.docx
```

---

## 文件结构

```
thesis-formatter/
├── SKILL.md                          # AI agent 调用入口
├── agents/openai.yaml                # Codex agent 配置
├── scripts/
│   ├── thesis_config.py              # 学校信息配置（首次必填）
│   ├── auto_format_docx.py           # 自动格式化主脚本
│   ├── generate_school_audit_inline_notes.py  # 行内批注审核
│   ├── generate_review_variants.py   # 批注版 + 自动修复版
│   ├── validate_format.py            # 格式验证
│   └── convert_to_three_line_table.py
├── references/
│   ├── format_specifications.md      # 格式参数参考
│   └── manual_review_checklist.md    # 人工复核清单
├── tests/                            # 单元测试
├── evals/evals.json                  # 评估用例
└── requirements.txt
```

---

## 依赖

- Python 3.8+
- `lxml >= 4.9`
- `python-docx >= 1.1`

---

## License

MIT

