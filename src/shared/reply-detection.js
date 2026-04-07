/**
 * Reply and forward separator detection for email threads.
 *
 * Extracted from commands.js to allow unit testing.
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

export function findTextSeparators(htmlBody) {
  const TAG_OR_GAP = "(?:\\s|<[^>]*>|&\\w+;|&#\\d+;|\\xA0)*";
  const fromRegex = new RegExp("\\b(De|From|Von|Van|Da|Fra)" + TAG_OR_GAP + ":", "gi");
  const confirmRegex = new RegExp(
    "\\b(Sent|Envoy(?:é|&eacute;|&#233;|e)|Enviado|Gesendet|Verzonden|Inviato" +
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

  // Detect Gmail/Apple Mail/Thunderbird inline attributions (past & present tense):
  // FR "a écrit", EN "wrote/writes", ES "escribió/escribe",
  // DE "schrieb/schreibt", NL "geschreven/schrijft", IT "scrisse/scrive"
  const wroteRegex =
    /\b(a\s+[eé]crit|wrot?e|writes|escribi[oó]|escribe|schrieb|schreibt|geschreven|schrijft|scrisse|scrive)\s*:/gi;
  const wrotePositions = [];
  let wroteMatch;
  // Matches preamble lines like "-------- Original Message --------"
  // or "-----Message d'origine-----" that often appear in a separate block
  // just before the "wrote:" attribution line (ProtonMail, Thunderbird).
  const preambleRegex =
    /<[^>]+>[\s-]*(?:Original Message|Message d'origine|Mensaje original|Ursprüngliche Nachricht|Origineel bericht|Messaggio originale)[\s-]*<\/[^>]+>\s*$/i;

  while ((wroteMatch = wroteRegex.exec(htmlBody)) !== null) {
    const lookback = htmlBody.substring(Math.max(0, wroteMatch.index - 500), wroteMatch.index);
    const blockTag = lookback.match(/.*(<(?:p|div|blockquote|li)\b[^>]*>)/is);
    let cutPos = blockTag
      ? wroteMatch.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : wroteMatch.index;

    // If a preamble line immediately precedes the block, include it in the cut.
    const before = htmlBody.substring(Math.max(0, cutPos - 300), cutPos);
    const preambleMatch = before.match(preambleRegex);
    if (preambleMatch) {
      cutPos = cutPos - before.length + before.indexOf(preambleMatch[0]);
    }

    wrotePositions.push(cutPos);
  }

  // Merge all strategies, sort by position, then deduplicate:
  // positions within 200 chars of each other belong to the same reply boundary.
  const all = [
    ...divPositions,
    ...borderPositions,
    ...hrPositions,
    ...textPositions,
    ...wrotePositions,
  ].sort((a, b) => a - b);

  const merged = [];
  for (const pos of all) {
    if (merged.length === 0 || pos - merged[merged.length - 1] >= 200) {
      merged.push(pos);
    }
  }
  return merged;
}
