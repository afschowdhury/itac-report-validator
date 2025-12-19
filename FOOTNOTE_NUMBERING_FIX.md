# Footnote Numbering Fix

## Issue
Footnote numbers displayed in the Web Link Validation UI were off by +1 from the actual footnote numbers in the Word document.

**Example:**
- UI showed: "Footnote 5", "Footnote 6"
- Word document: "Footnote 4", "Footnote 5"

## Root Cause

Word documents contain **special footnotes** in the XML that are not visible to users:
- **ID -1**: Separator footnote
- **ID 0**: Continuation separator  
- **ID 1**: Continuation notice

These are followed by user-visible footnotes:
- **ID 2**: User's "Footnote 1"
- **ID 3**: User's "Footnote 2"
- **ID 4**: User's "Footnote 3"
- **ID 5**: User's "Footnote 4"
- etc.

**The Problem:** We were displaying the XML `w:id` attribute directly (e.g., "Footnote 5") instead of the actual display number users see in Word (e.g., "Footnote 4").

## Solution

Updated `link_extractor.py` to:

1. **Filter out special footnotes** during extraction
   - Skip footnotes with `w:type` attribute (separator, continuationSeparator, continuationNotice)
   - Only process normal user footnotes

2. **Assign correct display numbers**
   - Counter starts at 1 for the first user footnote
   - Each user footnote gets a sequential display number
   - Store this as `display_number` in the footnotes dictionary

3. **Use display numbers in context**
   - When creating link context strings, use `display_number` instead of XML ID
   - Example: `f"Footnote {display_num}: {text}..."` instead of `f"Footnote {fn_id}: {text}..."`

## Code Changes

### `extract_footnotes_from_document()`
- Added `user_footnote_counter` to track sequential numbering
- Skip footnotes with `w:type` attribute (special footnotes)
- Store `display_number` in footnotes dictionary

### `extract_links_from_all_ars()`
- Use `footnote.get('display_number', fn_id)` when creating context
- Display correct footnote number to users

## Verification

### Before Fix:
```
Footnote 5: https://www.walmart.com/... (404)
Footnote 6: https://prosupplydirect.com/... (404)
```
(But these are actually Footnote 4 and 5 in Word)

### After Fix:
```
Footnote 4: https://www.walmart.com/... (404)  ✓ Correct!
Footnote 5: https://prosupplydirect.com/... (404)  ✓ Correct!
```

Now matches the actual footnote numbers in Word!

## Impact

✅ Footnote numbers in UI now match Word document exactly
✅ Users can easily find the correct footnote to fix
✅ No confusion between displayed and actual footnote numbers
✅ Maintains backward compatibility (works with docs without special footnotes)

## Files Modified
- `link_extractor.py`:
  - Updated `extract_footnotes_from_document()` 
  - Updated `extract_links_from_all_ars()`
  - Added display number tracking

