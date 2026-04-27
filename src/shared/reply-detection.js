/**
 * Reply and forward separator detection for email threads.
 *
 * Extracted from commands.js to allow unit testing.
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

    // If an underscore separator line (e.g. ________________________________)
    // immediately precedes the header block (with only empty elements between),
    // include it in the cut so the whole Outlook separator is removed.
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
      // Strip tags, HTML entities (&nbsp; etc.) and whitespace to check for
      // actual content between the underscore element and the De:/From: block.
      const between = preWindow
        .substring(lastUEnd)
        .replace(/<[^>]*>/g, "")
        .replace(/&[a-zA-Z]+;|&#\d+;/g, "")
        .replace(/[\s\xa0]/g, "");
      if (!between) {
        const movedCutPos = Math.max(0, cutPos - 400) + lastUIdx;
        // Only adopt the moved position if it still clears the deduplication
        // window; otherwise keep the original cutPos so the separator is not lost.
        if (positions.length === 0 || movedCutPos - positions[positions.length - 1] >= 200) {
          cutPos = movedCutPos;
        }
      }
    }

    // If a dash-separator line (e.g. -----Original Message-----)
    // immediately precedes the header block, include it in the cut.
    if (lastUIdx < 0) {
      const dashWindow = htmlBody.substring(Math.max(0, cutPos - 400), cutPos);
      const dashSepRe =
        /<[^>]+>\s*[-\u2010-\u2014]{3,}[\s\xa0]*[^-\u2010-\u2014\n\r<]{3,60}[\s\xa0]*[-\u2010-\u2014]{3,}\s*<\/[^>]+>/g;
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

  // Filter out empty divRplyFwdMsg elements — deeply nested forwards can
  // produce structural <div id="x_divRplyFwdMsg"></div> with no real content
  // after them (just closing tags).  These are not real reply boundaries.
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

  // Detect Gmail/Apple Mail/Thunderbird inline attributions (past & present tense):
  // FR "a écrit", EN "wrote/writes", ES "escribió/escribe",
  // DE "schrieb/schreibt", NL "geschreven/schrijft", IT "scrisse/scrive"
  // "a écrit" can be split across lines by a <br> tag when Outlook wraps long
  // attribution lines, so allow any mix of whitespace and HTML tags between
  // the auxiliary "a" and the past participle "écrit".
  // Gap between the attribution verb and the colon: plain whitespace plus the
  // common non-breaking-space entities.  Outlook emits `a écrit&nbsp;:` for
  // French attributions, which `\s*` alone would not match.
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
  // Matches preamble lines like "-------- Original Message --------"
  // or "-----Message d'origine-----" that often appear in a separate block
  // just before the "wrote:" attribution line (mimimail, Thunderbird).
  // No $ anchor: empty paragraphs between the preamble and the wrote: line
  // would break the end-of-string match.
  const preambleRegex =
    /<[^>]+>[\s\-\u2010-\u2014]*(?:Original Message|Message d'origine|Mensaje original|Ursprüngliche Nachricht|Origineel bericht|Messaggio originale|Forwarded Message|Message transféré|Mensaje reenviado|Weitergeleitete Nachricht|Doorgestuurd bericht|Messaggio inoltrato)[\s\-\u2010-\u2014]*<\/[^>]+>/i;

  while ((wroteMatch = wroteRegex.exec(htmlBody)) !== null) {
    // Skip attributions that are nested inside 2+ levels of plain-text quoting.
    // In Outlook, quoted replies use literal &gt; / > prefixes (not <blockquote>).
    // Two or more leading > on the same visual line means this "wrote:" is inside
    // a quote-within-a-quote, not a real reply boundary.
    // Note: <blockquote>-based nesting (Gmail, Apple Mail) is NOT filtered here
    // because each blockquote level introduces a real reply boundary.
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

    // Multi-line attributions: "wrote:" may be in a separate element from the
    // "On <date>, <name>" line.  Walk back through preceding sibling elements
    // that share the same quote prefix (> or &gt;) and contain attribution-like
    // text (starts with "On ", "Le ", "El ", "Am ", "Op ", "Il ", or a date).
    const attrLookStart = Math.max(0, cutPos - 600);
    const attrBefore = htmlBody.substring(attrLookStart, cutPos);
    // Find consecutive preceding block elements
    const prevBlockRe = /<(?:p|div)\b[^>]*>[\s\S]*?<\/(?:p|div)>/gi;
    const prevBlocks = [];
    let pbMatch;
    while ((pbMatch = prevBlockRe.exec(attrBefore)) !== null) {
      prevBlocks.push({ index: pbMatch.index, text: pbMatch[0] });
    }
    // Walk backwards from the last preceding block.
    // Two cases:
    //  - Quoted chain (Outlook plain-text replies): each line prefixed with >;
    //    walk back through consecutive quoted blocks until we find one starting
    //    with an attribution marker.
    //  - Unquoted Gmail/Apple-Mail attribution wrapped on 2 lines:
    //    "El lun, 24 mar 2025 a las 11:02, Alejandro Méndez"
    //    "<alejandro@...> escribió:"
    //    Allow a single jump to the IMMEDIATELY preceding block if it matches
    //    the attribution-start pattern, even if not quoted.
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
      // Stop walking back if this block doesn't look like part of the attribution
      if (!isQuoted) break;
    }

    // If a preamble line immediately precedes the block, include it in the cut.
    // Use a 500-char window and find the LAST preamble match to avoid picking up
    // an unrelated earlier one.
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

  // Standalone dash-separator lines like "-----Message d'origine-----" or
  // "-----Original Message-----".  These are unambiguous reply/forward
  // boundaries even when no standard email headers (De:/From:) follow.
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

  // Filter out dash-standalone separators (e.g. "------- Forwarded Message -------")
  // that appear inside the body of a reply already marked by divRplyFwdMsg or <hr>.
  // These are inline forwarded content, not a new reply boundary.
  const structuralPositions = [...divPositions, ...hrPositions];
  const dashStandalonePositions = dashStandalonePositionsRaw.filter((dp) => {
    return !structuralPositions.some((sp) => dp > sp && dp - sp < 1500);
  });

  // Filter out text-based positions (De:/From: + Subject/Objet:) that are
  // duplicates of a nearby structural separator (hr, divRplyFwdMsg, border-top,
  // dashStandalone).  A textPosition is considered a duplicate when:
  //  1. It falls within 1500 chars after a structural anchor, AND
  //  2. No other detected separator exists between the anchor and the textPosition.
  // Condition 2 prevents filtering genuine boundaries in short plain-text emails
  // where multiple replies can be close together (< 1500 chars apart).
  // Anchors with their dedup window.  divRplyFwdMsg is a short *envelope*
  // (De/Envoyé/Objet — ~400 chars); a De:/Subject text block more than ~600
  // chars after it is sibling content (a forwarded message that itself
  // contains plain-text headers), not a duplicate.  border-top:solid, <hr>
  // and dash-standalone wrap full reply bodies → 1500 chars.
  const anchorsWithWindow = [
    ...dashStandalonePositions.map((p) => ({ p, w: 1500 })),
    ...divPositions.map((p) => ({ p, w: 600 })),
    ...borderPositions.map((p) => ({ p, w: 1500 })),
    // <hr> is a thin boundary — only its immediately-following De:/Subject
    // cluster is the duplicate.  A text block >600 chars later is sibling
    // content (a forwarded message with its own plain-text headers).
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
      // If another separator exists between the anchor and this textPosition,
      // the textPosition belongs to a different reply — keep it.
      const hasIntermediate = allNonTextPositions.some((sp) => sp > p + 200 && sp < tp - 200);
      return !hasIntermediate;
    });
  });

  // Merge all strategies, sort by position, then deduplicate:
  // positions within 200 chars of each other belong to the same reply boundary.
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
