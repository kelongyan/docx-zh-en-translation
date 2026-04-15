# docx-zh-en-translation

A stable DOCX translation tool for Chinese-to-English workflows with structure preservation as the top priority.

## Overview

This project translates existing Chinese .docx documents into English while preserving Microsoft Word layout as much as possible. Instead of rebuilding a document from scratch, it edits conservative Office Open XML parts so that paragraphs, tables, headers, footers, comments, footnotes, and endnotes remain usable after translation.

It is intended for structured Word documents such as contracts, reports, proposals, manuals, and other files where formatting matters.

## Key Capabilities

- Preserve original .docx structure and generate a separate translated file
- Translate body text, table content, headers, footers, comments, footnotes, and endnotes
- Use a conservative strategy for high-risk Office objects
- Support configurable translation backends
- Repack and validate output documents after XML updates
- Use _en.docx as the default output suffix

## Repository Layout

    .
    |-- IEEE02.docx
    |-- IEEE02_en.docx
    |-- SKILL.md
    |-- AGENTS.md
    |-- README.md
    |-- evals/
    |   -- evals.json
    |-- references/
    |   -- xml-scope.md
    -- scripts/
        -- translate_docx.py

## Requirements

### Runtime

- Python 3.11 or newer
- A working translation backend
- Access to Office helper scripts for unpack, pack, and validate steps

### Backend

The implementation is backend-configurable. In the current codebase, translation can be performed through:

- a compatible API endpoint configured through environment variables, or
- a local claude CLI fallback when available

This README intentionally stays generic so the documentation does not need to change every time the backend setup evolves.

## Configuration

Set backend-related environment variables as needed:

    set LONGCAT_API_BASE=https://your-endpoint
    set LONGCAT_API_KEY=your_key
    set LONGCAT_MODEL=your_model_name

If the primary API path is unavailable, the script may fall back to a local CLI translator if supported by the environment.

## Usage

### Basic command

    python scripts/translate_docx.py <input.docx> [output.docx]

### Examples

    python scripts/translate_docx.py IEEE02.docx
    python scripts/translate_docx.py IEEE02.docx IEEE02_en.docx

### Output naming

- If no output path is provided, the default output is <original_name>_en.docx
- The output file is written beside the source document by default

## Processing Flow

### 1. Validate input

The script checks that the source file exists and is a valid .docx file.

### 2. Unpack DOCX

Helper scripts unpack the Office document into editable XML parts.

### 3. Select XML parts

The translator processes conservative text-bearing parts such as:

- word/document.xml
- word/styles.xml
- word/fontTable.xml
- word/comments.xml
- word/footnotes.xml
- word/endnotes.xml
- word/header*.xml
- word/footer*.xml

### 4. Extract text

Only visible and safe text nodes are collected. High-risk regions are skipped.

### 5. Translate

Chinese text is batched and translated through the configured backend.

### 6. Write back

Translated text is redistributed back into the original XML text nodes while minimizing structural disruption.

### 7. Repack and validate

The updated XML is repackaged into a new .docx, then validated to reduce the chance of producing a broken Word file.

## Translation Coverage

### Included

- Paragraph text
- Table cell text
- Headers and footers
- Comments
- Footnotes and endnotes

### Intentionally skipped

- Charts
- Equations and formulas
- Text inside images
- Field codes such as w:instrText
- Other high-risk drawing or embedded payload text

## Engineering Principles

- Preserve document integrity before chasing maximum extraction coverage
- Prefer minimal XML mutation over aggressive document reconstruction
- Keep batching explicit and configurable
- Centralize helper-script resolution and subprocess handling
- Use predictable naming and validation steps

## Validation

### Syntax check

    python -m py_compile scripts/translate_docx.py

### Translation test

    python scripts/translate_docx.py IEEE02.docx

Expected default output:

    IEEE02_en.docx

### Manual review checklist

After generation, verify at least the following:

- The output file opens normally in Word
- Main Chinese content has been translated into English
- Tables remain intact
- Headers, footers, comments, footnotes, and endnotes are preserved
- Charts and other skipped objects are not damaged

## FAQ

### Why does the project not translate every possible object inside a DOCX?

Some Office objects are too fragile to edit safely at the XML level. This project intentionally prioritizes document integrity over maximum text extraction.

### Why is the default output name in English?

Using _en.docx avoids Windows terminal and tooling issues related to non-ASCII filenames while still making the output purpose clear.

### Can the translation backend change later?

Yes. The implementation supports backend configuration, so the documentation remains stable even if the underlying provider evolves.

## Security Notes

- Do not commit real API keys
- Do not commit client documents without permission
- Use controlled environments for sensitive contracts or internal materials

## References

- scripts/translate_docx.py
- SKILL.md
- 
eferences/xml-scope.md
- vals/evals.json

# docx-zh-en-translation
