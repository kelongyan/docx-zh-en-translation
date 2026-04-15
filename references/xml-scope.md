# XML scope for DOCX translation skill

## Included XML parts

- `word/document.xml`
- `word/header*.xml`
- `word/footer*.xml`
- `word/comments.xml`
- `word/footnotes.xml`
- `word/endnotes.xml`

## Included content

- visible paragraph text in `<w:t>`
- text inside table cells
- header/footer text
- comment body text
- footnote/endnote text
- text inside `w:txbxContent` when it is standard WordprocessingML text and not part of chart/drawing payloads

## Excluded content

- `w:instrText`
- chart XML parts
- formula / math objects
- drawing payload text where structure is not ordinary WordprocessingML paragraphs
- image OCR / embedded raster text
- relationship files, style definitions, numbering definitions

## Safe-write principle

Prefer preserving paragraph/run structure and replace only text-bearing nodes.

If a paragraph has multiple plain text runs, translation may be redistributed across those runs proportionally.
If redistribution is not safe, collapse visible text into the first writable run and blank later writable runs in the same logical segment.

## Known limitations

- English expansion can change line wrapping and page breaks
- highly formatted inline text may not preserve run-level emphasis perfectly
- TOC / field-driven content is intentionally skipped
