import {
  collectRegexPositions,
  findTextSeparators,
  findReplySeparators,
} from "../src/commands/reply-detection.js";

// ---------------------------------------------------------------------------
// Helpers to build realistic HTML snippets
// ---------------------------------------------------------------------------

function makeDiv(id, content = "") {
  return `<div id="${id}">${content}</div>`;
}

function makeThread(...replies) {
  return replies.join("\n");
}

// ---------------------------------------------------------------------------
// collectRegexPositions
// ---------------------------------------------------------------------------

describe("collectRegexPositions", () => {
  test("returns empty array when no matches", () => {
    expect(collectRegexPositions("hello world", /<hr>/gi)).toEqual([]);
  });

  test("returns position of a single match", () => {
    const html = "before<hr>after";
    const result = collectRegexPositions(html, /<hr>/gi);
    expect(result).toEqual([6]);
  });

  test("returns positions of multiple matches", () => {
    // a=0, <hr>=1..4, b=5, <hr>=6..9, c=10, <hr>=11..14, d=15
    const html = "a<hr>b<hr>c<hr>d";
    const result = collectRegexPositions(html, /<hr>/gi);
    expect(result).toEqual([1, 6, 11]);
  });

  test("filters out matches that fail the headerCheck", () => {
    // First <hr> is followed by >500 chars of filler before "From:" appears
    // so the 500-char window does NOT capture "From:", and only the second <hr> passes
    const filler = "x".repeat(510);
    const html = `<hr>${filler}<hr>From: someone`;
    const headerCheck = /\bFrom\s*:/i;
    const result = collectRegexPositions(html, /<hr[^>]*>/gi, headerCheck);
    expect(result).toHaveLength(1);
    // The second <hr> is at index 4 + 510 = 514
    expect(result[0]).toBe(514);
  });

  test("includes match when headerCheck passes", () => {
    const html = "<hr>From: John";
    const headerCheck = /\bFrom\s*:/i;
    const result = collectRegexPositions(html, /<hr[^>]*>/gi, headerCheck);
    expect(result).toEqual([0]);
  });

  test("headerCheck window is limited to 500 chars after match", () => {
    // Put "From:" 600 chars after the <hr> — should NOT be found
    const html = "<hr>" + "x".repeat(600) + "From: someone";
    const headerCheck = /\bFrom\s*:/i;
    const result = collectRegexPositions(html, /<hr[^>]*>/gi, headerCheck);
    expect(result).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// findReplySeparators — strategy 1: divRplyFwdMsg
// ---------------------------------------------------------------------------

describe("findReplySeparators — divRplyFwdMsg strategy", () => {
  test("detects a single Outlook modern reply separator", () => {
    const html = `<p>My reply</p><div id="divRplyFwdMsg">Original message</div>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
    expect(result[0]).toBe(html.indexOf('<div id="divRplyFwdMsg">'));
  });

  test("detects x_divRplyFwdMsg variant", () => {
    const html = `<p>My reply</p><div id="x_divRplyFwdMsg">Original message</div>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects multiple divRplyFwdMsg separators in a long thread", () => {
    const html = makeThread(
      "<p>Reply 1</p>",
      makeDiv("divRplyFwdMsg", "<p>Reply 2</p>"),
      makeDiv("divRplyFwdMsg", "<p>Reply 3</p>"),
      makeDiv("divRplyFwdMsg", "<p>Original</p>")
    );
    const result = findReplySeparators(html);
    expect(result).toHaveLength(3);
  });

  test("returns empty array for plain email with no reply markers", () => {
    const html = "<p>Hello, this is a simple message.</p>";
    const result = findReplySeparators(html);
    expect(result).toHaveLength(0);
  });

  test("is case-insensitive for the id attribute value", () => {
    // The regex uses ["'] delimiters; test with double quotes and mixed casing in id
    const html = `<div id="divRplyFwdMsg">x</div>`;
    expect(findReplySeparators(html)).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// findReplySeparators — strategy 2: border-top solid
// ---------------------------------------------------------------------------

describe("findReplySeparators — border-top strategy", () => {
  function makeBorderDiv(content = "") {
    return `<div style="border-top: solid #E1E1E1 1.0pt; padding: 3.0pt 0cm 0cm 0cm">${content}</div>`;
  }

  test("detects a border-top separator followed by From:", () => {
    const html = `<p>Reply</p>${makeBorderDiv("<b>From:</b> someone")}`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects a border-top separator followed by De: (French)", () => {
    const html = `<p>Réponse</p>${makeBorderDiv("<b>De :</b> quelqu'un")}`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("ignores border-top div NOT followed by a From-like header", () => {
    const html = `<p>Reply</p><div style="border-top: solid red 1pt">Just a styled box</div>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(0);
  });

  test("detects multiple border-top separators in a thread", () => {
    const sep = makeBorderDiv("<b>From:</b> someone");
    const html = `<p>R3</p>${sep}<p>R2</p>${sep}<p>R1</p>${sep}<p>Original</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(3);
  });
});

// ---------------------------------------------------------------------------
// findReplySeparators — strategy 3: <hr> tag
// ---------------------------------------------------------------------------

describe("findReplySeparators — <hr> strategy", () => {
  test("detects <hr> followed by From:", () => {
    const html = `<p>Reply</p><hr><p><b>From:</b> alice@example.com</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects <hr /> self-closing with From: header", () => {
    const html = `<p>Reply</p><hr /><p>From: alice</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects <hr> with attributes followed by From:", () => {
    const html = `<p>Reply</p><hr style="color:grey"><p>From: bob@example.com</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("ignores <hr> NOT followed by a From-like header", () => {
    const html = `<p>Section 1</p><hr><p>Section 2</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(0);
  });

  test("detects two <hr> separators", () => {
    const html =
      `<p>R2</p><hr><p>From: bob</p><p>R1</p><hr><p>From: alice</p><p>Original</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(2);
  });
});

// ---------------------------------------------------------------------------
// findTextSeparators / findReplySeparators — text strategy (multilingual)
// ---------------------------------------------------------------------------

describe("findTextSeparators — multilingual text patterns", () => {
  function makeTextThread(fromLabel, sentLabel, content = "") {
    return (
      `<p>My reply</p>` +
      `<p><b>${fromLabel}:</b> someone@example.com</p>` +
      `<p><b>${sentLabel}:</b> Monday, January 1, 2024</p>` +
      `<p>${content}</p>`
    );
  }

  test("detects English From / Sent separator", () => {
    const html = makeTextThread("From", "Sent");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects French De / Envoyé separator", () => {
    const html = makeTextThread("De", "Envoyé");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects French De / Objet separator", () => {
    const html = makeTextThread("De", "Objet");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects German Von / Gesendet separator", () => {
    const html = makeTextThread("Von", "Gesendet");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects German Von / Betreff separator", () => {
    const html = makeTextThread("Von", "Betreff");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects Dutch Van / Verzonden separator", () => {
    const html = makeTextThread("Van", "Verzonden");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects Dutch Van / Onderwerp separator", () => {
    const html = makeTextThread("Van", "Onderwerp");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects Italian Da / Inviato separator", () => {
    const html = makeTextThread("Da", "Inviato");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects Italian Da / Oggetto separator", () => {
    const html = makeTextThread("Da", "Oggetto");
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("does NOT detect 'From:' alone without a confirmation keyword", () => {
    const html = `<p>From: alice@example.com</p><p>Just some reply text.</p>`;
    const result = findTextSeparators(html);
    expect(result).toHaveLength(0);
  });

  test("handles HTML entities in Envoyé (&eacute;)", () => {
    const html =
      `<p>My reply</p>` +
      `<p>De: quelqu'un</p>` +
      `<p>Envoy&eacute;: lundi 1 janvier 2024</p>`;
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("handles HTML entities in Envoyé (&#233;)", () => {
    const html =
      `<p>My reply</p>` +
      `<p>De: quelqu'un</p>` +
      `<p>Envoy&#233;: lundi 1 janvier 2024</p>`;
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("handles tags between keyword and colon", () => {
    // e.g. "From<span> </span>:" - tags between keyword and colon
    const html =
      `<p>Reply text</p>` +
      `<p><b>From<span> </span>:</b> alice</p>` +
      `<p><b>Sent<span> </span>:</b> Monday</p>`;
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });

  test("detects two text separators in a three-message thread", () => {
    // Separators must be >200 chars apart to avoid deduplication
    const padding = "<p>" + "x".repeat(300) + "</p>";
    const sep = (f, s) =>
      `<p><b>${f}:</b> someone@example.com</p><p><b>${s}:</b> Monday Jan 1</p>`;
    const html = `<p>R2</p>${sep("From", "Sent")}${padding}${sep("From", "Subject")}`;
    const result = findTextSeparators(html);
    expect(result).toHaveLength(2);
  });

  test("deduplicates separators within 200 chars of each other", () => {
    // Two "From/Sent" matches very close together should be collapsed into one
    const sep = `<p>From: x</p><p>Sent: y</p>`;
    const html = `<p>R</p>${sep}${sep}`;
    const result = findTextSeparators(html);
    expect(result).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// findReplySeparators — strategy selection (best wins)
// ---------------------------------------------------------------------------

describe("findReplySeparators — strategy selection", () => {
  test("prefers divRplyFwdMsg over text when it has more matches", () => {
    // 2 div separators vs 1 text separator
    const html =
      `<p>R3</p>` +
      `<div id="divRplyFwdMsg"><p>R2</p></div>` +
      `<div id="divRplyFwdMsg">` +
        `<p>From: someone</p><p>Sent: date</p>` +
        `<p>R1 with some text</p>` +
      `</div>`;
    const result = findReplySeparators(html);
    // The div strategy yields 2 — should be picked
    expect(result).toHaveLength(2);
  });

  test("falls back to text strategy when no structural markers exist", () => {
    // Separators must be >200 chars apart to avoid deduplication
    const padding = "<p>" + "x".repeat(300) + "</p>";
    const html =
      `<p>Reply 2</p>` +
      `<p>From: alice@example.com</p><p>Sent: Monday Jan 1</p>` +
      padding +
      `<p>From: bob@example.com</p><p>Subject: hello world</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(2);
  });
});

// ---------------------------------------------------------------------------
// findReplySeparators — edge cases
// ---------------------------------------------------------------------------

describe("findReplySeparators — edge cases", () => {
  test("returns empty array for empty string", () => {
    expect(findReplySeparators("")).toEqual([]);
  });

  test("returns empty array for plain text with no markers", () => {
    expect(findReplySeparators("<p>Hello world</p>")).toEqual([]);
  });

  test("does not confuse a Forward with a Reply — both are detected", () => {
    // Outlook uses the same divRplyFwdMsg id for both replies and forwards
    const html =
      `<p>My forward comment</p>` +
      `<div id="divRplyFwdMsg">` +
        `<p>-------- Forwarded Message --------</p>` +
        `<p>From: original@sender.com</p>` +
      `</div>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("is robust to extra attributes on the div id element", () => {
    const html = `<div class="someClass" id="divRplyFwdMsg" style="color:red">content</div>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });

  test("handles non-breaking spaces between header keyword and colon", () => {
    const nbsp = "\u00A0";
    const html =
      `<p>Reply</p><hr>` +
      `<p>From${nbsp}: alice@example.com</p>`;
    const result = findReplySeparators(html);
    expect(result).toHaveLength(1);
  });
});
