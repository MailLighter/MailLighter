/**
 * Pure functions for detecting reply and forward separators in email HTML bodies.
 * Extracted here to be independently testable without Office.js dependencies.
 */

/**
 * Collects positions of all regex matches in an HTML body.
 * Optionally validates each match against a header pattern.
 *
 * @param {string} htmlBody - The HTML content to search
 * @param {RegExp} regex - The regex to match (must have the 'g' flag)
 * @param {RegExp} [headerCheck] - Optional regex to validate the 500 chars after each match
 * @returns {number[]} Array of match positions (indices)
 */
export function collectRegexPositions(htmlBody, regex, headerCheck) {
  const positions = [];
  let match;
  while ((match = regex.exec(htmlBody)) !== null) {
    if (headerCheck) {
      const after = htmlBody.substring(match.index, match.index + 500);
      if (!headerCheck.test(after)) continue;
    }
    positions.push(match.index);
  }
  return positions;
}

/**
 * Detects reply chain separators using multilingual text patterns.
 * Looks for "From:" / "De:" / "Von:" etc. followed by "Sent:" / "Envoyé:" etc.
 *
 * Supported languages: English, French, German, Dutch, Italian, Danish/Norwegian.
 *
 * @param {string} htmlBody - The HTML content to search
 * @returns {number[]} Array of positions where reply separators start
 */
export function findTextSeparators(htmlBody) {
  const TAG_OR_GAP = "(?:\\s|<[^>]*>|&\\w+;|&#\\d+;|\\xA0)*";
  const fromRegex = new RegExp(
    "\\b(De|From|Von|Van|Da|Fra)" + TAG_OR_GAP + ":",
    "gi"
  );
  const confirmRegex = new RegExp(
    "\\b(Sent|Envoy(?:é|&eacute;|&#233;|e)|Gesendet|Verzonden|Inviato" +
      "|Objet|Subject|Betreff|Onderwerp|Oggetto)" +
      TAG_OR_GAP +
      ":",
    "i"
  );

  const positions = [];
  let match;
  while ((match = fromRegex.exec(htmlBody)) !== null) {
    const after = htmlBody.substring(match.index, match.index + 1500);
    if (!confirmRegex.test(after)) continue;
    const lookback = htmlBody.substring(Math.max(0, match.index - 500), match.index);
    const blockTag = lookback.match(/.*(<(?:p|div|tr|li)\b[^>]*>)/is);
    const cutPos = blockTag
      ? match.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : match.index;
    if (positions.length > 0 && cutPos - positions[positions.length - 1] < 200) continue;
    positions.push(cutPos);
  }
  return positions;
}

/**
 * Detects all reply/forward separators in an HTML email body using 4 strategies:
 * 1. Div with id="divRplyFwdMsg" or "x_divRplyFwdMsg" (modern Outlook)
 * 2. Div with border-top:solid style + header keyword check (Outlook web)
 * 3. <hr> tag + header keyword check (generic)
 * 4. Multilingual text patterns (fallback)
 *
 * Returns the strategy that yields the most separators.
 *
 * @param {string} htmlBody - The HTML content to search
 * @returns {number[]} Array of positions where reply/forward separators start
 */
export function findReplySeparators(htmlBody) {
  const headerPattern = /\b(From|De|Von|Da|Van|Fra)\s*(&nbsp;|\xA0)?\s*:/i;

  const divPositions = collectRegexPositions(
    htmlBody,
    /<div[^>]*\bid\s*=\s*["'](?:x_)*divRplyFwdMsg["'][^>]*>/gi
  );

  const borderPositions = collectRegexPositions(
    htmlBody,
    /<div[^>]*border-top\s*:\s*solid\s[^>]*>/gi,
    headerPattern
  );

  const hrPositions = collectRegexPositions(htmlBody, /<hr[^>]*>/gi, headerPattern);

  const textPositions = findTextSeparators(htmlBody);

  let best = divPositions;
  if (borderPositions.length > best.length) best = borderPositions;
  if (hrPositions.length > best.length) best = hrPositions;
  if (textPositions.length > best.length) best = textPositions;
  return best;
}
