---
name: md-to-branded-docx
description: "Convert Markdown files to professionally formatted Word documents (.docx) with XD.AI corporate branding. Applies company letterhead (logo in header), standard fonts (Causten family), brand colors (#3313E2 blue), and paragraph styles. Use when: (1) Converting markdown notes/reports to branded Word documents, (2) Creating client-ready documents from markdown, (3) Generating professional reports with XD.AI branding, (4) Transforming technical documentation to corporate format."
---

# Markdown to Branded DOCX Converter

Convert Markdown files to Word documents with XD.AI corporate branding, including logo, fonts, colors, and styling.

## Quick Start

```bash
python /path/to/skill/scripts/md_to_docx.py input.md output.docx
```

## What Gets Applied

The conversion applies these XD.AI brand elements from the template:

| Element | Specification |
|---------|---------------|
| Logo | XD.AI logo in header (left-aligned) |
| Footer | "Page X of Y" (right-aligned, Causten 10pt) |
| Body Font | Causten, 11pt |
| Heading 1 | Causten Bold, 20pt, brand blue (#3313E2) |
| Heading 2 | Causten Medium, 16pt, black |
| Heading 3 | Causten Medium, 14pt, black |
| Page | A4, 25mm margins |
| Code | Consolas, 10pt, gray background |

## Supported Markdown Elements

- **Headings**: `# H1` through `###### H6`
- **Paragraphs**: Regular text
- **Bold**: `**text**` or `__text__`
- **Italic**: `*text*` or `_text_`
- **Bold+Italic**: `***text***`
- **Inline code**: `` `code` ``
- **Code blocks**: Triple backticks with optional language
- **Bullet lists**: `- item` or `* item` (nested supported)
- **Numbered lists**: `1. item` (nested supported)
- **Blockquotes**: `> quote`
- **Horizontal rules**: `---` or `***`
- **Links**: `[text](url)` (rendered as underlined blue text)

## Usage Examples

### Basic conversion
```bash
python scripts/md_to_docx.py report.md report.docx
```

### With cover page
```bash
python scripts/md_to_docx.py report.md report.docx --cover
```
Adds a full-page brand blue (#3313E2) cover with the XD.AI logo, the document title (from first `# H1`), and the current month/year.

### With table of contents
```bash
python scripts/md_to_docx.py report.md report.docx --toc
```
Inserts a "Contents" section with a Word TOC field that covers H1-H3. Open the file in Word and right-click the field to update it.

### Cover + TOC (most common for long reports)
```bash
python scripts/md_to_docx.py report.md report.docx --cover --toc
```

### H1 section breaks
Automatic. Every `# Heading 1` after the first one gets a page break before it, creating clean section separation.

### Programmatic usage
```python
from md_to_docx import MarkdownToDocx

converter = MarkdownToDocx(cover=True, toc=True)
converter.convert(markdown_string, 'output.docx')
```

## Bundled Assets

- `assets/template.dotx` - XD.AI Word template with all styles pre-configured
- `assets/logo.png` - XD.AI logo (used in header)

## Brand Reference

See [references/brand-styles.md](references/brand-styles.md) for complete typography, color, and spacing specifications extracted from the template.

## Troubleshooting

**Font substitution**: If Causten fonts aren't installed on the viewing system, Word will substitute. The document embeds font names but not font files.

**List numbering**: Complex nested lists use the template's built-in numbering definitions. If lists appear incorrectly, verify the template's numbering.xml is intact.

**Logo not showing**: Ensure `assets/logo.png` exists and the header references in the template are preserved.
