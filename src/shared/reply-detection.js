/**
 * Reply/Forward separator detection for email HTML bodies.
 *
 * Supports four strategies (Outlook HTML markers, CSS border dividers, <hr> tags,
 * text-based From/Sent headers) plus dashed separators used by Thunderbird, Apple
 * Mail, Gmail mobile and plain-text Outlook (e.g. "-------- Original Message --------"
 * or "-----Message d'origine-----").
 */

/**
 * Collect all positions in htmlBody where regex matches, optionally requiring
 * a headerCheck pattern to match within the next 500 characters.
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
 * Detect Outlook-style text separators: a "From:" field followed within 1500
 * characters by a "Sent:" / "Subject:" / equivalent field. Handles multiple
 * languages and HTML entity encoding.
 */
export function findTextSeparators(htmlBody) {
  const TAG_OR_GAP = "(?:\\s|<[^>]*>|&\\w+;|&#\\d+;|\\xA0)*";
  const fromRegex = new RegExp(
    "\\b(De|From|Von|Van|Da|Fra)" + TAG_OR_GAP + ":",
    "gi"
  );
  const confirmRegex = new RegExp(
    "\\b(Sent|Envoy(?:é|&eacute;|&#233;|e)|Enviado(?:\\s+el)?|Gesendet|Verzonden|Inviato" +
      "|Objet|Subject|Asunto|Betreff|Onderwerp|Oggetto)" +
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
 * Detect dashed separators of the form:
 *   -------- Original Message --------     (Thunderbird, Gmail mobile, Apple Mail)
 *   > -------- Original Message --------   (quoted-style, &gt; in HTML)
 *   -----Message d'origine-----            (French Outlook plain-text)
 *   -----Ursprüngliche Nachricht-----      (German)
 *   … and equivalent labels in other supported languages.
 *
 * The apostrophe in "d'origine" may be a literal apostrophe, &#39;, &apos;,
 * or a Unicode right single quotation mark (\u2019).
 * Accented characters may be literal UTF-8 or HTML numeric/named entities.
 */
export function findDashedSeparators(htmlBody) {
  const APOS = "(?:'|&#39;|&apos;|\u2019)";
  const E = "(?:é|&eacute;|&#233;)";
  const U_UML = "(?:ü|&uuml;|&#252;)";

  const LABELS = [
    // English
    "Original\\s+Message",
    "Original\\s+Appointment",
    "Forwarded\\s+[Mm]essage",
    // French
    "Message\\s+d" + APOS + "origine",
    "Message\\s+transf" + E + "r" + E,
    // German
    "Urspr" + U_UML + "ngliche\\s+Nachricht",
    "Weitergeleitete\\s+Nachricht",
    // Spanish
    "Mensaje\\s+original",
    "Mensaje\\s+reenviado",
    // Italian
    "Messaggio\\s+originale",
    "Messaggio\\s+inoltrato",
    // Dutch
    "Doorgestuurd\\s+bericht",
    "Oorspronkelijk\\s+bericht",
  ].join("|");

  // Optional leading &gt; (HTML-encoded >) followed by optional spaces, then
  // 3+ dashes, the label, 3+ dashes.
  const dashedSepRegex = new RegExp(
    "(?:&gt;\\s*)?-{3,}\\s*(?:" + LABELS + ")\\s*-{3,}",
    "gi"
  );

  const positions = [];
  let match;
  while ((match = dashedSepRegex.exec(htmlBody)) !== null) {
    // Walk back to the start of the nearest enclosing block element so the
    // cut point is clean (same logic as findTextSeparators).
    const lookback = htmlBody.substring(Math.max(0, match.index - 200), match.index);
    const blockTag = lookback.match(/.*(<(?:p|div|tr|li|blockquote)\b[^>]*>)/is);
    const cutPos = blockTag
      ? match.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : match.index;
    if (positions.length > 0 && cutPos - positions[positions.length - 1] < 200) continue;
    positions.push(cutPos);
  }
  return positions;
}

/**
 * Find all reply/forward separator positions in an HTML email body using the
 * best-matching strategy among five candidates. Returns an array of integer
 * character positions, one per separator, in document order.
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

  const dashedPositions = findDashedSeparators(htmlBody);

  let best = divPositions;
  if (borderPositions.length > best.length) best = borderPositions;
  if (hrPositions.length > best.length) best = hrPositions;
  if (textPositions.length > best.length) best = textPositions;
  if (dashedPositions.length > best.length) best = dashedPositions;
  return best;
}
