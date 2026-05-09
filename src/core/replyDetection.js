/**
 * Reply and forward separator detection for email threads.
 *
 * Logic kept identical to the previous src/shared/reply-detection.js module
 * (moved here as part of the Phase 1a refactor). No behavioural change.
 */

/**
 * Collect all positions (indices) in `htmlBody` where `regex` matches.
 *
 * @param {string} htmlBody - The HTML string to search.
 * @param {RegExp} regex - **Must have the global (`g`) flag set.** Without it
 *   the `while` loop will not advance `lastIndex` and will loop infinitely.
 * @param {RegExp|null} [headerCheck] - Optional secondary regex applied to the
 *   200-char window after each match; the position is only recorded when this
 *   check passes.
 * @returns {number[]} Sorted array of character positions.
 */
export function collectRegexPositions(htmlBody, regex, headerCheck) {
  const positions = [];
  let match;
  while ((match = regex.exec(htmlBody)) !== null) {
    if (headerCheck) {
      const after = htmlBody.substring(match.index, match.index + 2000);
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
    const after = htmlBody.substring(match.index, match.index + 3000);
    if (!confirmRegex.test(after)) continue;
    const lookback = htmlBody.substring(Math.max(0, match.index - 500), match.index);
    const blockTag = lookback.match(/.*(<(?:p|div|tr|li)\b[^>]*>)/is);
    let cutPos = blockTag
      ? match.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : match.index;

    const preWindow = htmlBody.substring(Math.max(0, cutPos - 400), cutPos);
    const underscoreRe = /<[^>]+>\s*_{10,}\s*<\/[^>]+>/g;
    let lastUIdx = -1;
    let lastUEnd = 0;
    let uTest;
    underscoreRe.lastIndex = 0;
    while ((uTest = underscoreRe.exec(preWindow)) !== null) {
      lastUIdx = uTest.index;
      lastUEnd = uTest.index + uTest[0].length;
    }
    if (lastUIdx >= 0) {
      const between = preWindow
        .substring(lastUEnd)
        .replace(/<[^>]*>/g, "")
        .replace(/&[a-zA-Z]+;|&#\d+;/g, "")
        .replace(/[\s\xa0]/g, "");
      if (!between) {
        const movedCutPos = Math.max(0, cutPos - 400) + lastUIdx;
        if (positions.length === 0 || movedCutPos - positions[positions.length - 1] >= 200) {
          cutPos = movedCutPos;
        }
      }
    }

    if (lastUIdx < 0) {
      const dashWindow = htmlBody.substring(Math.max(0, cutPos - 400), cutPos);
      const dashSepRe =
        /<[^>]+>\s*[-‐-—]{3,}[\s\xa0]*[^-‐-—\n\r<]{3,60}[\s\xa0]*[-‐-—]{3,}\s*<\/[^>]+>/g;
      let lastDIdx = -1;
      let lastDEnd = 0;
      let dTest;
      while ((dTest = dashSepRe.exec(dashWindow)) !== null) {
        lastDIdx = dTest.index;
        lastDEnd = dTest.index + dTest[0].length;
      }
      if (lastDIdx >= 0) {
        const betweenDash = dashWindow
          .substring(lastDEnd)
          .replace(/<[^>]*>/g, "")
          .replace(/&[a-zA-Z]+;|&#\d+;/g, "")
          .replace(/[\s\xa0]/g, "");
        if (!betweenDash) {
          const movedCutPos = Math.max(0, cutPos - 400) + lastDIdx;
          if (positions.length === 0 || movedCutPos - positions[positions.length - 1] >= 200) {
            cutPos = movedCutPos;
          }
        }
      }
    }

    if (positions.length > 0 && cutPos - positions[positions.length - 1] < 200) continue;
    positions.push(cutPos);
  }
  return positions;
}

export function findReplySeparators(htmlBody) {
  const headerPattern = /\b(From|De|Von|Da|Van|Fra)\s*(&nbsp;|\xA0)?\s*:/i;

  const divPositionsRaw = collectRegexPositions(
    htmlBody,
    /<div[^>]*\bid\s*=\s*["'](?:x_)*divRplyFwdMsg["'][^>]*>/gi
  );

  const divPositions = divPositionsRaw.filter((pos) => {
    const after = htmlBody.substring(pos, Math.min(htmlBody.length, pos + 500));
    const textContent = after
      .replace(/<[^>]*>/g, "")
      .replace(/&[a-zA-Z]+;|&#\d+;/g, "")
      .replace(/[\s\xa0]/g, "");
    return textContent.length > 0;
  });

  const borderPositions = collectRegexPositions(
    htmlBody,
    /<div[^>]*border-top\s*:[^;]*\bsolid\b[^>]*>/gi,
    headerPattern
  );

  const hrPositions = collectRegexPositions(htmlBody, /<hr[^>]*>/gi, headerPattern);

  const textPositions = findTextSeparators(htmlBody);

  const ATTR_GAP = "(?:\\s|&nbsp;|&#160;|&#xA0;)*";
  const wroteRegex = new RegExp(
    "\\b(a(?:\\s|<[^>]*>)+[eé]crit" +
      "|wrot?e|writes|escribi[oó]|escribe|schrieb|schreibt|geschreven|schrijft|scrisse|scrive)" +
      ATTR_GAP +
      ":",
    "gi"
  );
  const wrotePositions = [];
  let wroteMatch;
  const preambleRegex =
    /<[^>]+>[\s\-‐-—]*(?:Original Message|Message d'origine|Mensaje original|Ursprüngliche Nachricht|Origineel bericht|Messaggio originale|Forwarded Message|Message transféré|Mensaje reenviado|Weitergeleitete Nachricht|Doorgestuurd bericht|Messaggio inoltrato)[\s\-‐-—]*<\/[^>]+>/i;

  while ((wroteMatch = wroteRegex.exec(htmlBody)) !== null) {
    const lineWindow = htmlBody.substring(Math.max(0, wroteMatch.index - 300), wroteMatch.index);
    const lineText = lineWindow.replace(/<[^>]*>/g, "\n");
    const lastLine =
      lineText
        .split("\n")
        .filter((l) => l.trim())
        .pop() || "";
    if (/^[\s\xa0]*(?:(?:&gt;|>)[\s\xa0]*){2,}/.test(lastLine)) continue;

    const lookback = htmlBody.substring(Math.max(0, wroteMatch.index - 500), wroteMatch.index);
    const blockTag = lookback.match(/.*(<(?:p|div|blockquote|li)\b[^>]*>)/is);
    let cutPos = blockTag
      ? wroteMatch.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : wroteMatch.index;

    const attrLookStart = Math.max(0, cutPos - 600);
    const attrBefore = htmlBody.substring(attrLookStart, cutPos);
    const prevBlockRe = /<(?:p|div)\b[^>]*>[\s\S]*?<\/(?:p|div)>/gi;
    const prevBlocks = [];
    let pbMatch;
    while ((pbMatch = prevBlockRe.exec(attrBefore)) !== null) {
      prevBlocks.push({ index: pbMatch.index, text: pbMatch[0] });
    }
    for (let bi = prevBlocks.length - 1; bi >= 0; bi--) {
      const pb = prevBlocks[bi];
      const plainText = pb.text
        .replace(/<[^>]*>/g, "")
        .replace(/&nbsp;/gi, " ")
        .replace(/&gt;/g, ">")
        .trim();
      const isQuoted = /^>/.test(plainText);
      const unquoted = plainText.replace(/^(?:>\s*)+/, "").trim();
      const isAttribution = /^(?:On |Le |El |Am |Op |Il |\d{1,2}[\s/.-])/.test(unquoted);
      const isImmediate = bi === prevBlocks.length - 1;
      if (isAttribution && (isQuoted || isImmediate)) {
        const candidatePos = attrLookStart + pb.index;
        if (
          wrotePositions.length === 0 ||
          candidatePos - wrotePositions[wrotePositions.length - 1] >= 200
        ) {
          cutPos = candidatePos;
        }
        break;
      }
      if (!isQuoted) break;
    }

    const lookbackStart = Math.max(0, cutPos - 500);
    const before = htmlBody.substring(lookbackStart, cutPos);
    const preambleGlobal = new RegExp(preambleRegex.source, "gi");
    let preambleMatch = null;
    let pm;
    while ((pm = preambleGlobal.exec(before)) !== null) {
      preambleMatch = pm;
    }
    if (preambleMatch && preambleMatch.index >= before.length - 400) {
      cutPos = lookbackStart + preambleMatch.index;
    }

    wrotePositions.push(cutPos);
  }

  const dashSepStandaloneRe = new RegExp(
    "<[^>]+>\\s*[-\\u2010-\\u2014]{3,}[\\s\\xa0]*(?:" +
      "Original Message|Message d'origine|Mensaje original" +
      "|Urspr(?:ü|&uuml;|&#252;)ngliche Nachricht|Origineel bericht|Messaggio originale" +
      "|Forwarded Message|Message transf(?:é|&eacute;|&#233;)r(?:é|&eacute;|&#233;)" +
      "|Mensaje reenviado|Weitergeleitete Nachricht|Doorgestuurd bericht|Messaggio inoltrato" +
      ")[\\s\\xa0]*[-\\u2010-\\u2014]{3,}\\s*<\\/[^>]+>",
    "gi"
  );
  const dashStandalonePositionsRaw = collectRegexPositions(htmlBody, dashSepStandaloneRe);

  const structuralPositions = [...divPositions, ...hrPositions];
  const dashStandalonePositions = dashStandalonePositionsRaw.filter((dp) => {
    return !structuralPositions.some((sp) => dp > sp && dp - sp < 1500);
  });

  const anchorsWithWindow = [
    ...dashStandalonePositions.map((p) => ({ p, w: 1500 })),
    ...divPositions.map((p) => ({ p, w: 600 })),
    ...borderPositions.map((p) => ({ p, w: 1500 })),
    ...hrPositions.map((p) => ({ p, w: 600 })),
  ];
  const allNonTextPositions = [
    ...dashStandalonePositions,
    ...divPositions,
    ...borderPositions,
    ...hrPositions,
    ...wrotePositions,
  ].sort((a, b) => a - b);
  const filteredTextPositions = textPositions.filter((tp) => {
    return !anchorsWithWindow.some(({ p, w }) => {
      if (tp <= p || tp - p >= w) return false;
      const hasIntermediate = allNonTextPositions.some((sp) => sp > p + 200 && sp < tp - 200);
      return !hasIntermediate;
    });
  });

  const all = [
    ...divPositions,
    ...borderPositions,
    ...hrPositions,
    ...filteredTextPositions,
    ...wrotePositions,
    ...dashStandalonePositions,
  ].sort((a, b) => a - b);

  const merged = [];
  for (const pos of all) {
    if (merged.length === 0 || pos - merged[merged.length - 1] >= 200) {
      merged.push(pos);
    }
  }
  return merged;
}
