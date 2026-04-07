# Security Audit - MailLighter Outlook Plugin

**Date:** 2026-04-01
**Version:** 0.0.1
**Overall Risk Level:** LOW

---

## Executive Summary

MailLighter is an Office Add-in for Microsoft Outlook that cleans emails before sending (removes images, attachments, trims reply threads, keeps selection only). The plugin processes all data locally within the Outlook client with no external network communication, no data storage, and no authentication requirements.

**Key finding:** The HTML sanitization function was incomplete, allowing potential XSS vectors (event handlers, dangerous tags, `javascript:` URIs) to pass through. This has been remediated with targeted regex rules while preserving Outlook-specific HTML compatibility.

---

## Architecture Security Assessment

### Data Flow

```
User action in Outlook Ribbon
    |
    v
JavaScript Command Handler (commands.js)
    |
    v
Office.js API (getAsync / setAsync)
    |
    v
Current email item (in-memory, local)
    |
    v
[NO external network calls]
```

### Permission Model

| Aspect | Value |
|--------|-------|
| Permission level | `ReadWriteItem` |
| Scope | Current email item only |
| Mailbox access | No |
| Contact access | No |
| Calendar access | No |
| Credential handling | None |
| External API calls | None |

---

## Vulnerabilities Found

### V1 - Incomplete HTML Sanitization (MEDIUM) - REMEDIATED

- **File:** `src/shared/office-helpers.js`
- **Issue:** `sanitizeSelectionHtml()` only removed `<script>` tags and HTML comments. Event handlers (`onerror`, `onload`), dangerous tags (`<iframe>`, `<embed>`, `<object>`, `<applet>`, `<form>`), and `javascript:`/`data:text/html` URIs were not filtered.
- **Impact:** Mitigated by Office.js sandbox, but a malicious email selection could inject active HTML content when reinserted via `setBodyAsync`.
- **Fix applied:** Added targeted regex rules to strip event handlers, dangerous tags, and dangerous URI schemes while preserving legitimate Outlook HTML (tables, `mso-*` styles, `<o:p>` namespace tags, `cid:` image links).

### V2 - Verbose Error Messages (LOW) - REMEDIATED

- **File:** `src/commands/commands.js`
- **Issue:** Raw Office.js API error messages were displayed to users in notification toasts, potentially revealing implementation details.
- **Fix applied:** Error details are now logged to `console.error()` for debugging while user-facing notifications show only generic localized messages.

### V3 - Development CORS Wildcard (LOW) - ACCEPTED

- **File:** `webpack.config.js`
- **Issue:** Dev server uses `Access-Control-Allow-Origin: *`.
- **Decision:** Accepted risk. Dev server runs on `localhost:3000` only, not publicly accessible. Restricting the origin risks breaking the development workflow across different Outlook clients (Old Outlook, New Outlook, OWA) which use varying origins.

### V4 - Regex-Based HTML Parsing (LOW) - ACCEPTED

- **File:** `src/commands/commands.js`
- **Issue:** Reply separator detection uses regex patterns on HTML body instead of DOM parsing.
- **Decision:** Accepted risk. The regex patterns do not have nested quantifiers vulnerable to ReDoS. Input is limited to the current email body (not arbitrary external input). DOM parsing would require a dependency like DOMPurify, which strips Outlook-specific HTML.

---

## Positive Security Properties

| Property | Status |
|----------|--------|
| Zero network communication | No `fetch`, `XMLHttpRequest`, `WebSocket`, or HTTP calls |
| Minimal permissions | `ReadWriteItem` only (current email) |
| No data storage | No `localStorage`, `sessionStorage`, `cookies`, or databases |
| No secrets in code | No API keys, tokens, or credentials; `.env` in `.gitignore` |
| Source maps disabled in production | `devtool: false` in production webpack config |
| HTTPS enforced | All manifest URLs use HTTPS |
| ESLint Office plugin | `eslint-plugin-office-addins` with recommended rules |
| No dangerous APIs | No `eval()`, `Function()`, `innerHTML`, or `dangerouslySetInnerHTML` |
| HTML escaping | `escapeHtml()` covers all 5 special characters |
| Non-destructive | Original emails are never modified |
| Minimal dependencies | Only 2 production dependencies (`core-js`, `regenerator-runtime`) |

---

## Compliance

| Standard | Status |
|----------|--------|
| GDPR / Privacy by Design | Compliant - local processing only, no data collection |
| OWASP Top 10 | Compliant - HTML sanitization covers XSS vectors |
| Microsoft Office Add-in Security | Compliant - minimal permissions, HTTPS, ESLint |
| No data exfiltration | Compliant - zero external network calls |

---

## Recommendations for Future Development

1. **Add automated tests** for `sanitizeSelectionHtml()` with OWASP XSS payloads
2. **Add `npm audit`** to CI/CD pipeline when one is established
3. **Consider DOMPurify** only after automated test coverage is in place, with a permissive config preserving Outlook HTML
4. **Create SECURITY.md** with vulnerability reporting policy if the plugin is distributed publicly
