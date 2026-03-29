import {
  collectRegexPositions,
  findTextSeparators,
  findDashedSeparators,
  findReplySeparators,
} from "./reply-detection";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Wrap content in a minimal HTML email skeleton. */
function email(...paragraphs) {
  return `<html><body>${paragraphs.join("")}</body></html>`;
}

function p(text) {
  return `<p>${text}</p>`;
}

/**
 * Paragraph whose total HTML length exceeds 200 chars, which is the
 * deduplication threshold in findTextSeparators / findDashedSeparators.
 * Without this, two separators whose HTML `<p>` tags are < 200 chars apart
 * would be collapsed into one by the deduplication guard.
 */
function longP(label) {
  // Approximately 220-char pad ensures the `<p>` tag distance between
  // consecutive separators is > 200 chars (threshold = separator html ~41 + body > 159).
  const pad =
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
    "Sed ut perspiciatis unde omnis iste natus error sit voluptatem " +
    "accusantium doloremque laudantium, totam rem aperiam.";
  return `<p>${label}: ${pad}</p>`;
}

// ---------------------------------------------------------------------------
// collectRegexPositions
// ---------------------------------------------------------------------------

describe("collectRegexPositions", () => {
  test("returns empty array when no match", () => {
    expect(collectRegexPositions("<p>hello</p>", /foo/gi)).toEqual([]);
  });

  test("returns position of single match", () => {
    const html = "<p>hello</p><hr><p>world</p>";
    const positions = collectRegexPositions(html, /<hr>/gi);
    expect(positions).toEqual([html.indexOf("<hr>")]);
  });

  test("returns multiple positions in order", () => {
    const html = "<hr><p>x</p><hr><p>y</p><hr>";
    const positions = collectRegexPositions(html, /<hr>/gi);
    expect(positions).toHaveLength(3);
    expect(positions[0]).toBeLessThan(positions[1]);
    expect(positions[1]).toBeLessThan(positions[2]);
  });

  test("headerCheck filters out matches without nearby header", () => {
    // First <hr> is followed by 600+ chars of filler — "From:" is NOT within 500 chars.
    // Second <hr> is immediately followed by "From:", so it passes the check.
    const filler = "x".repeat(600);
    const html = `<hr><p>${filler}</p><hr><p>From: a@b.com</p>`;
    const headerCheck = /\bFrom\s*:/i;
    const positions = collectRegexPositions(html, /<hr>/gi, headerCheck);
    expect(positions).toHaveLength(1);
    expect(html.substring(positions[0]).startsWith("<hr>")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// findTextSeparators
// ---------------------------------------------------------------------------

describe("findTextSeparators", () => {
  test("detects English From/Subject pair", () => {
    const html = email(
      p("My reply"),
      p("From: sender@example.com"),
      p("Subject: Re: test")
    );
    expect(findTextSeparators(html)).toHaveLength(1);
  });

  test("detects French De/Objet pair", () => {
    const html = email(
      p("Ma réponse"),
      p("De : expediteur@example.com"),
      p("Objet : Re: test")
    );
    expect(findTextSeparators(html)).toHaveLength(1);
  });

  test("detects German Von/Betreff pair", () => {
    const html = email(
      p("Meine Antwort"),
      p("Von: absender@example.com"),
      p("Betreff: Aw: test")
    );
    expect(findTextSeparators(html)).toHaveLength(1);
  });

  test("detects Envoyé with HTML entity &eacute;", () => {
    const html = email(
      p("Ma réponse"),
      p("De : expediteur@example.com"),
      p("Envoy&eacute; : lundi 1 janvier")
    );
    expect(findTextSeparators(html)).toHaveLength(1);
  });

  test("returns empty when From without matching Sent/Subject", () => {
    const html = email(p("From: orphan@example.com"));
    expect(findTextSeparators(html)).toHaveLength(0);
  });

  test("detects multiple separators in a thread", () => {
    // Each reply body must exceed the 200-char deduplication threshold.
    const block = email(
      longP("Reply 3"),
      p("From: c@c.com"), p("Sent: Mon"), p("Subject: test"),
      longP("Reply 2"),
      p("From: b@b.com"), p("Sent: Sun"), p("Subject: test"),
      longP("Reply 1"),
      p("From: a@a.com"), p("Sent: Sat"), p("Subject: test"),
      p("Original")
    );
    const positions = findTextSeparators(block);
    expect(positions.length).toBeGreaterThanOrEqual(3);
  });
});

// ---------------------------------------------------------------------------
// findDashedSeparators — pattern 2: Thunderbird/mobile style
// ---------------------------------------------------------------------------

describe("findDashedSeparators — Thunderbird/Apple Mail style (Original Message)", () => {
  test("detects plain dashed separator without leading >", () => {
    const html = email(
      p("My reply"),
      p("-------- Original Message --------"),
      p("From: sender@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects separator with HTML-encoded leading > (&gt;)", () => {
    const html = email(
      p("My reply"),
      p("&gt; -------- Original Message --------"),
      p("&gt; From: sender@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects separator with varying dash counts", () => {
    const html = email(
      p("Reply"),
      p("--- Original Message ---"),
      p("From: a@b.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects 'Forwarded Message' separator", () => {
    const html = email(
      p("See below"),
      p("-------- Forwarded Message --------"),
      p("From: a@b.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects 'Forwarded message' (lowercase m)", () => {
    const html = email(
      p("FYI"),
      p("-------- Forwarded message --------"),
      p("From: a@b.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects 3 separators in a deep thread", () => {
    const thread =
      longP("Reply C") +
      p("-------- Original Message --------") +
      longP("Reply B") +
      p("-------- Original Message --------") +
      longP("Reply A") +
      p("-------- Original Message --------") +
      p("Original");
    expect(findDashedSeparators(email(thread))).toHaveLength(3);
  });

  test("returns empty array when no dashed separator present", () => {
    const html = email(p("Just a plain message without any separator."));
    expect(findDashedSeparators(html)).toHaveLength(0);
  });

  test("does not match plain dash lines (no label)", () => {
    const html = email(p("--------------------"), p("Not a separator"));
    expect(findDashedSeparators(html)).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// findDashedSeparators — pattern 3: Outlook plain-text French
// ---------------------------------------------------------------------------

describe("findDashedSeparators — Outlook plain-text French (Message d'origine)", () => {
  test("detects with literal apostrophe", () => {
    const html = email(
      p("Ma réponse"),
      p("-----Message d'origine-----"),
      p("De : expediteur@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects with HTML numeric entity &#39;", () => {
    const html = email(
      p("Ma réponse"),
      p("-----Message d&#39;origine-----"),
      p("De : expediteur@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects with &apos; entity", () => {
    const html = email(
      p("Ma réponse"),
      p("-----Message d&apos;origine-----"),
      p("De : expediteur@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects with Unicode right single quote (\u2019)", () => {
    const html = email(
      p("Ma réponse"),
      p("-----Message d\u2019origine-----"),
      p("De : expediteur@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects with space between dashes and label", () => {
    const html = email(
      p("Ma réponse"),
      p("----- Message d'origine -----"),
      p("De : expediteur@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects 'Message transféré' with literal accent", () => {
    const html = email(
      p("Voir ci-dessous"),
      p("-----Message transféré-----"),
      p("De : expediteur@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects 'Message transféré' with &eacute; entities", () => {
    const html = email(
      p("Voir ci-dessous"),
      p("-----Message transf&eacute;r&eacute;-----"),
      p("De : expediteur@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects multiple French separators in thread", () => {
    const thread =
      longP("Réponse 2") +
      p("-----Message d'origine-----") +
      longP("Réponse 1") +
      p("-----Message d'origine-----") +
      p("Message initial");
    expect(findDashedSeparators(email(thread))).toHaveLength(2);
  });
});

// ---------------------------------------------------------------------------
// findDashedSeparators — other language variants
// ---------------------------------------------------------------------------

describe("findDashedSeparators — other languages", () => {
  test("detects German 'Ursprüngliche Nachricht' (literal ü)", () => {
    const html = email(
      p("Meine Antwort"),
      p("-------- Ursprüngliche Nachricht --------"),
      p("Von: absender@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects German 'Ursprüngliche Nachricht' with &uuml;", () => {
    const html = email(
      p("Meine Antwort"),
      p("-------- Urspr&uuml;ngliche Nachricht --------"),
      p("Von: absender@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects German 'Weitergeleitete Nachricht'", () => {
    const html = email(
      p("Zur Info"),
      p("-------- Weitergeleitete Nachricht --------"),
      p("Von: absender@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects Spanish 'Mensaje original'", () => {
    const html = email(
      p("Mi respuesta"),
      p("-------- Mensaje original --------"),
      p("De: remitente@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects Italian 'Messaggio originale'", () => {
    const html = email(
      p("La mia risposta"),
      p("-------- Messaggio originale --------"),
      p("Da: mittente@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });

  test("detects Dutch 'Doorgestuurd bericht'", () => {
    const html = email(
      p("Zie hieronder"),
      p("-------- Doorgestuurd bericht --------"),
      p("Van: afzender@example.com")
    );
    expect(findDashedSeparators(html)).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// findReplySeparators — integration: dashed separators are picked up
// ---------------------------------------------------------------------------

describe("findReplySeparators — integration with dashed separators", () => {
  test("finds dashed separator when no Outlook HTML markers are present", () => {
    const html = email(
      p("My reply"),
      p("-------- Original Message --------"),
      p("From: sender@example.com")
    );
    expect(findReplySeparators(html)).toHaveLength(1);
  });

  test("finds French dashed separator", () => {
    const html = email(
      p("Ma réponse"),
      p("-----Message d'origine-----"),
      p("De : expediteur@example.com")
    );
    expect(findReplySeparators(html)).toHaveLength(1);
  });

  test("prefers divRplyFwdMsg count over dashed when higher", () => {
    // Build an email with 2 Outlook div markers and only 1 dashed separator.
    const html =
      '<div id="divRplyFwdMsg">reply1</div>' +
      p("-------- Original Message --------") +
      '<div id="divRplyFwdMsg">reply2</div>';
    const positions = findReplySeparators(html);
    // The div strategy wins (2 > 1), so both div positions are returned.
    expect(positions).toHaveLength(2);
  });

  test("returns empty array for a clean compose message", () => {
    const html = email(p("Bonjour,"), p("Mon message ici."));
    expect(findReplySeparators(html)).toHaveLength(0);
  });

  test("finds 3 dashed separators in a deep thread", () => {
    const html = email(
      longP("Reply C"),
      p("-------- Original Message --------"),
      longP("Reply B"),
      p("-------- Original Message --------"),
      longP("Reply A"),
      p("-------- Original Message --------"),
      p("Original")
    );
    expect(findReplySeparators(html)).toHaveLength(3);
  });
});

// ---------------------------------------------------------------------------
// findReplySeparators — existing strategies still work after refactoring
// ---------------------------------------------------------------------------

describe("findReplySeparators — existing Outlook HTML strategies", () => {
  test("detects divRplyFwdMsg", () => {
    const html =
      '<p>My reply</p>' +
      '<div id="divRplyFwdMsg"><p>From: a@b.com</p></div>';
    expect(findReplySeparators(html)).toHaveLength(1);
  });

  test("detects x_divRplyFwdMsg variant", () => {
    const html =
      '<p>My reply</p>' +
      '<div id="x_divRplyFwdMsg"><p>From: a@b.com</p></div>';
    expect(findReplySeparators(html)).toHaveLength(1);
  });

  test("detects <hr> with nearby From:", () => {
    const html =
      "<p>My reply</p>" +
      "<hr>" +
      "<p>From: sender@example.com</p>" +
      "<p>Sent: Monday</p>";
    expect(findReplySeparators(html)).toHaveLength(1);
  });

  test("detects text-based From/Subject separator", () => {
    const html = email(
      p("My reply"),
      p("From: sender@example.com"),
      p("Subject: Re: test")
    );
    expect(findReplySeparators(html)).toHaveLength(1);
  });
});
