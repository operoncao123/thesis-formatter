# Installation

## 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/phd-thesis-formatter.git
cd phd-thesis-formatter
```

## 2. Install dependencies

```bash
pip install -r requirements.txt
```

## 3. Configure your institution

Edit `scripts/thesis_config.py` and fill in your institution's details:

```python
SCHOOL_NAME = 'YOUR_UNIVERSITY'          # e.g. '北京大学'
EVEN_PAGE_HEADER = SCHOOL_NAME + '博士学位论文'
MIN_PAGES_BEFORE_REFERENCES = 100
REQUIRED_POST_REF_SECTIONS = ['致谢', '原创性声明', '使用授权说明']
```

## 4. Using with Claude Code (CLI)

Copy or symlink this skill directory into your Claude Code skills folder:

```bash
# macOS / Linux
mkdir -p ~/.claude/skills
ln -s "$(pwd)" ~/.claude/skills/phd-thesis-formatter
```

Then in Claude Code, you can invoke the skill with:

```
$phd-thesis-formatter
```

or just describe what you need and Claude Code will trigger it automatically.

## 5. Using with OpenAI Codex

Copy or symlink this skill directory into your Codex skills folder:

```bash
mkdir -p ~/.codex/skills
ln -s "$(pwd)" ~/.codex/skills/phd-thesis-formatter
```

The `agents/openai.yaml` file registers the skill with the Codex agent interface.

## 6. Run directly (without an AI agent)

```bash
# Full auto-format pass
python scripts/auto_format_docx.py thesis.docx --output thesis_formatted.docx --report report.txt

# Inline red-note audit
python scripts/generate_school_audit_inline_notes.py thesis.docx --output thesis_audit.docx

# Inline review + conservative auto-fix
python scripts/generate_review_variants.py thesis.docx \
  --inline-output thesis_review.docx \
  --fixed-output thesis_fixed.docx \
  --report review_report.txt

# Validate format only
python scripts/validate_format.py thesis.docx
```
