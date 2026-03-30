# YOUR_UNIVERSITY博士学位论文格式化 Skill 说明

## 定位

`phd-thesis-formatter` 现在被定义为“半自动格式化 + 人工复核”的技能。

它适合处理接近定稿的 `.docx` 论文，帮助完成：

- 页面设置归一化
- 基础字体归一化
- 普通表格转三线表
- 对明确编号标题补 Word 标题样式
- 基础结构检查和格式报告

它**不应**被当作“任何博士论文都能一键完全合规”的全自动工具。

## 当前可自动完成

- A4 纸张与页边距设置
- 页眉/页脚距离设置
- 基础字体统计与部分字体补全
- 三线表边框转换
- 运行格式验证脚本并输出通过项/告警项

## 当前必须人工复核

- 标题层级是否判断正确
- 奇偶页页眉与分节页码
- 目录域更新
- 匿名送审版脱敏
- 参考文献从 APA/混杂格式到 Nature 风格的准确转换
- 跨页续表、分图、复杂公式与上下标

## 主要文件

- `SKILL.md`: skill 主体说明
- `references/format_specifications.md`: 北大格式参数参考
- `references/manual_review_checklist.md`: 高风险人工复核项
- `scripts/convert_to_three_line_table.py`: 三线表转换脚本
- `scripts/validate_format.py`: 基础验证脚本

## 推荐使用方式

```bash
python scripts/auto_format_docx.py input.docx --output output.docx --report format_report.txt
```

如需逐步检查 XML，再使用：

```bash
DOCX_SKILL_DIR="${CODEX_HOME:-$HOME/.codex}/skills/docx"
python "$DOCX_SKILL_DIR/scripts/office/unpack.py" input.docx unpacked/
python scripts/convert_to_three_line_table.py unpacked/word/document.xml
python "$DOCX_SKILL_DIR/scripts/office/pack.py" unpacked/ output.docx --original input.docx
python scripts/validate_format.py output.docx
```

## 输出要求

输出文件应至少包括：

1. `[原文件名]_formatted.docx`
2. 一份报告，分成：
   - 已自动修复
   - 需要人工复核
   - 未处理/不能安全自动处理

## 评估标准

只有当验证脚本通过，并且人工复核清单中的关键项也确认无误后，才可以对用户说“基本符合要求”。
