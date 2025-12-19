# Link Categorization Update

## Issue
Links returning **403 Forbidden** were being marked as "Broken", but they are actually accessible via web browsers (just blocked for automated requests due to bot detection, popups, or CAPTCHA requirements).

## Solution
Refined the link status categorization to be more accurate for real-world usage.

## Updated Categorization Logic

### 🟢 Working (Green)
- **HTTP 200-299**: Successful responses
- No redirects, fast response, valid SSL

### 🟡 Warning (Yellow)
- **HTTP 401 Unauthorized**: Requires authentication (resource exists)
- **HTTP 403 Forbidden**: Bot detection, CAPTCHA, or cookie requirements (likely accessible to real users)
- **HTTP 400-499** (except 404): Client-side issues but resource may exist
  - 400 Bad Request
  - 405 Method Not Allowed
  - 429 Too Many Requests
  - etc.
- **HTTP 300-399**: Redirects
- **Slow response**: > 5 seconds
- **SSL issues**: Invalid or expired certificates

### 🔴 Broken (Red)
- **HTTP 404 Not Found**: Resource definitively doesn't exist
- **HTTP 500-599**: Server errors (server down or malfunctioning)
- **Network errors**: Timeout, DNS failure, connection errors

## Updated Suggestion Messages

### 403 Forbidden
**Before:**
> "Access forbidden (403). You may not have permission to access this resource."

**After:**
> "Access forbidden (403). This is often due to bot detection or cookie requirements. The page may be accessible in a regular web browser with popups/CAPTCHA enabled."

### Other 4xx Errors
Added specific messages for:
- **400**: Bad request - URL syntax issues
- **405**: Method not allowed - page exists but doesn't accept GET
- **429**: Rate limiting - try again later
- **Other 4xx**: Generic client error message

## Impact

### Before Update
- All 403 responses → **Broken** (Red) ❌
- Users confused why "broken" links work in browser

### After Update
- 403 responses → **Warning** (Yellow) ⚠️
- Clear explanation that it's likely bot detection
- Only truly broken links (404, 5xx) marked as **Broken**

## Testing

To verify the changes work correctly:

1. Upload a document with links
2. Check links that return 403
3. Verify they appear in **Warnings** section (yellow badge)
4. Read the suggestion message explaining bot detection

## Files Modified
- `link_validator.py`:
  - Updated `categorize_link_status()` function
  - Enhanced `get_error_suggestion()` with specific 4xx messages

## Benefits
✅ More accurate categorization  
✅ Reduces false positives for "broken" links  
✅ Better user experience with clearer explanations  
✅ Helps users understand which links truly need fixing (404s) vs which might just need browser access (403s)

