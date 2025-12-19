# Link Extraction Fix - Footnote Support

## Problem
The link extraction feature was not finding links in DOCX documents because:
- Links were stored in Word document **footnotes**
- The original implementation only scanned paragraph text and tables
- Footnotes were completely ignored

## Solution Implemented

### 1. Updated `link_extractor.py`
Added comprehensive footnote extraction support:

**New Functions:**
- `extract_footnotes_from_document(doc)` - Extracts all footnotes from the document
  - Accesses footnotes via document relationships (`doc.part.rels`)
  - Parses footnotes XML blob using `parse_xml()`
  - Extracts both hyperlinks and plain text URLs from each footnote
  - Returns dict mapping footnote IDs to their content and links

- `find_footnote_references_in_blocks(ar_blocks)` - Finds footnote references in AR blocks
  - Scans paragraphs and tables for `w:footnoteReference` elements
  - Returns list of footnote IDs referenced in the AR

**Updated Function:**
- `extract_links_from_all_ars()` - Now accepts `Document` object instead of path
  - Extracts all footnotes once (efficient)
  - For each AR, finds its footnote references
  - Adds links from referenced footnotes to the AR's link list
  - Marks location as `'footnote'` for proper display

### 2. Updated `document_extractor.py`
Modified to pass Document object through the extraction pipeline:

- `build_outputs()` - Added `doc` parameter
- `extract_itac_report()` - Passes `doc` object to `build_outputs()`
- Link extraction now has access to full document structure

## Results

### Before Fix:
```
No web links found in Assessment Recommendations
```

### After Fix:
```
INFO:link_extractor:Found 41 footnote elements in document
INFO:link_extractor:Extracted 41 footnotes with content
INFO:link_extractor:Found 7 link(s) in AR_02 (including footnotes)
INFO:link_extractor:Found 8 link(s) in AR_03 (including footnotes)
INFO:link_extractor:Found 6 link(s) in AR_04 (including footnotes)
INFO:link_extractor:Found 9 link(s) in AR_05 (including footnotes)
INFO:link_extractor:Found 1 link(s) in AR_06 (including footnotes)
INFO:link_extractor:Found 5 link(s) in AR_07 (including footnotes)

Total links found: 36
From footnotes: 36
From paragraphs: 0
From tables: 0
```

## Technical Details

### Footnote Access Method
python-docx doesn't expose footnotes directly, so we:
1. Access the footnotes part via document relationships
2. Parse the raw XML blob: `parse_xml(footnotes_part.blob)`
3. Find all `<w:footnote>` elements
4. Extract hyperlinks (`<w:hyperlink>`) and plain text URLs
5. Map footnote IDs to their references in AR blocks

### Link Location Types
- `'footnote'` - Link found in a document footnote
- `'paragraph'` - Link found in paragraph text
- `'table'` - Link found in a table cell

## Testing

Tested with: `docs/report4/LS2521 - Draft.docx`
- Document contains 41 footnotes
- 36 links extracted from footnotes referenced in ARs
- Links span across 6 different Assessment Recommendations

## UI Impact

The Web Link Validation section now:
- Shows all links found in ARs (including those from footnotes)
- Displays location type (footnote, paragraph, or table)
- Validates all URLs (working, warning, broken)
- Provides actionable fix suggestions

## Files Modified
1. `/link_extractor.py` - Added footnote extraction logic
2. `/document_extractor.py` - Pass Document object for footnote access
3. `/app.py` - Already updated (no changes needed)
4. `/templates/comparison.html` - Already updated (no changes needed)

## No Breaking Changes
- Existing functionality preserved
- Backward compatible (works with docs without footnotes)
- Same data structure and API

