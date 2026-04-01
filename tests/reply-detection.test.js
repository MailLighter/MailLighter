const {
  findReplySeparators,
  findTextSeparators,
} = require("../src/shared/reply-detection");

// ===========================================================================
// findTextSeparators — header-block detection (De/From + confirmation word)
// ===========================================================================
describe("findTextSeparators", () => {
  // -----------------------------------------------------------------------
  // English
  // -----------------------------------------------------------------------
  test("detects English Outlook header (From + Sent + Subject)", () => {
    const html =
      "<p>Hello</p>" +
      "<p>From: John <john@test.com><br>Sent: Monday<br>Subject: Test</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(1);
  });

  test("detects multiple English headers", () => {
    const html =
      "<p>Body</p>" +
      "<p>From: A<br>Subject: S1</p>" +
      "<p>" + "x".repeat(300) + "</p>" +
      "<p>From: B<br>Subject: S2</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(2);
  });

  // -----------------------------------------------------------------------
  // French
  // -----------------------------------------------------------------------
  test("detects French Outlook header (De + Envoyé + Objet)", () => {
    const html =
      "<p>Bonjour</p>" +
      "<p>De : Marie <marie@test.fr><br>Envoyé : lundi<br>Objet : Test</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Spanish
  // -----------------------------------------------------------------------
  test("detects Spanish Outlook header (De + Enviado + Asunto)", () => {
    const html =
      "<p>Hola</p>" +
      "<p>De: Carmen <carmen@test.es><br>Enviado el: lunes<br>Asunto: Proyecto</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(1);
  });

  test("detects Spanish Thunderbird forward (De + Asunto)", () => {
    const html =
      "<p>FYI</p>" +
      "<p>De: nagios@monitoring.es<br>Fecha: martes<br>Para: sysadmins@test.es<br>Asunto: [NAGIOS] CRITICO</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // German
  // -----------------------------------------------------------------------
  test("detects German Outlook header (Von + Gesendet + Betreff)", () => {
    const html =
      "<p>Hallo</p>" +
      "<p>Von: Hans <hans@test.de><br>Gesendet: Montag<br>Betreff: Test</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Dutch
  // -----------------------------------------------------------------------
  test("detects Dutch Outlook header (Van + Verzonden + Onderwerp)", () => {
    const html =
      "<p>Hallo</p>" +
      "<p>Van: Jan <jan@test.nl><br>Verzonden: maandag<br>Onderwerp: Test</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Italian
  // -----------------------------------------------------------------------
  test("detects Italian Outlook header (Da + Inviato + Oggetto)", () => {
    const html =
      "<p>Ciao</p>" +
      "<p>Da: Marco <marco@test.it><br>Inviato: lunedì<br>Oggetto: Test</p>";
    const seps = findTextSeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Edge cases
  // -----------------------------------------------------------------------
  test("returns empty array when no separators found", () => {
    const html = "<p>Just a normal email with no quotes.</p>";
    expect(findTextSeparators(html)).toEqual([]);
  });

  test("does not match De/From without a confirmation word", () => {
    const html = "<p>Message from De: someone without Subject or Sent.</p>";
    expect(findTextSeparators(html)).toEqual([]);
  });
});

// ===========================================================================
// findReplySeparators — all strategies combined
// ===========================================================================
describe("findReplySeparators", () => {
  // -----------------------------------------------------------------------
  // Strategy 1: divRplyFwdMsg (Outlook standard)
  // -----------------------------------------------------------------------
  test("detects Outlook divRplyFwdMsg", () => {
    const html =
      "<p>My reply</p>" +
      '<div id="divRplyFwdMsg"><p>From: A<br>Sent: Mon<br>Subject: S</p></div>';
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(1);
  });

  test("detects x_divRplyFwdMsg (prefixed variant)", () => {
    const html =
      "<p>My reply</p>" +
      '<div id="x_divRplyFwdMsg"><p>From: A</p></div>';
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Strategy 2: border-top: solid (Outlook separator line)
  // -----------------------------------------------------------------------
  test("detects border-top solid div with From header", () => {
    const html =
      "<p>Reply</p>" +
      '<div style="border-top: solid #B5C4DF 1.0pt">' +
      "<p>From: Someone<br>Sent: Mon</p></div>";
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Strategy 3: <hr> with From header
  // -----------------------------------------------------------------------
  test("detects <hr> followed by From header", () => {
    const html =
      "<p>Reply</p>" +
      "<hr>" +
      "<p>From: Someone<br>Sent: Monday</p>";
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Strategy 4: text-based (findTextSeparators — tested above in detail)
  // -----------------------------------------------------------------------
  test("detects -----Original Message----- pattern via text strategy", () => {
    const html =
      "<p>Reply</p>" +
      "<p>-----Original Message-----</p>" +
      "<p>From: John<br>Sent: Mon<br>Subject: Test</p>";
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(1);
  });

  // -----------------------------------------------------------------------
  // Strategy 5: wrote/a écrit attributions
  // -----------------------------------------------------------------------
  describe("Gmail/Apple Mail/Thunderbird attributions", () => {
    // English
    test("detects 'wrote:' (EN past)", () => {
      const html = "<p>On Mon, Mar 24, 2025, John &lt;john@test.com&gt; wrote:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    test("detects 'writes:' (EN present, Thunderbird/mutt)", () => {
      const html = "<p>John &lt;john@test.com&gt; writes:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    // French
    test("detects 'a écrit :' (FR)", () => {
      const html = "<p>Le lun. 24 mars 2025 à 09:14, Marie &lt;m@test.fr&gt; a écrit :</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    test("detects 'a ecrit :' (FR without accent)", () => {
      const html = "<p>Le lun. 24 mars, Marie a ecrit :</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    // Spanish
    test("detects 'escribió:' (ES past, Gmail)", () => {
      const html = "<p>El lun, 24 mar 2025 a las 11:02, Alejandro &lt;a@test.es&gt; escribió:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    test("detects 'escribe:' (ES present, Thunderbird)", () => {
      const html = "<p>Paula Giménez &lt;paula@test.com&gt; escribe:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    // German
    test("detects 'schrieb:' (DE past, direct colon)", () => {
      // Note: Gmail DE format is "Am ... schrieb Name <email>:" where ":"
      // follows the email, not "schrieb" directly. This tests the simpler
      // Thunderbird DE case where "schrieb:" is immediately followed by colon.
      const html = "<p>Hans &lt;h@test.de&gt; schrieb:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    test("detects 'schreibt:' (DE present, Thunderbird)", () => {
      const html = "<p>Hans &lt;h@test.de&gt; schreibt:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    // Dutch
    test("detects 'geschreven:' (NL past)", () => {
      const html = "<p>Op ma 24 mrt 2025, Jan &lt;j@test.nl&gt; geschreven:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    test("detects 'schrijft:' (NL present, Thunderbird)", () => {
      const html = "<p>Jan &lt;j@test.nl&gt; schrijft:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    // Italian
    test("detects 'scrisse:' (IT past)", () => {
      const html = "<p>Il 24 mar 2025, Marco &lt;m@test.it&gt; scrisse:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });

    test("detects 'scrive:' (IT present, Thunderbird)", () => {
      const html = "<p>Marco &lt;m@test.it&gt; scrive:</p>";
      const seps = findReplySeparators(html);
      expect(seps.length).toBe(1);
    });
  });

  // -----------------------------------------------------------------------
  // Multi-reply threads — best strategy wins
  // -----------------------------------------------------------------------
  test("picks strategy with the most separators", () => {
    const html =
      "<p>My reply</p>" +
      "<p>On Mon, John wrote:</p>" +
      "<blockquote><p>First reply</p></blockquote>" +
      "<p>" + "x".repeat(200) + "</p>" +
      "<p>On Tue, Jane wrote:</p>" +
      "<blockquote><p>Second reply</p></blockquote>" +
      "<p>" + "x".repeat(200) + "</p>" +
      "<p>On Wed, Bob wrote:</p>" +
      "<blockquote><p>Third reply</p></blockquote>";
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(3);
  });

  // -----------------------------------------------------------------------
  // Spanish full Outlook thread (real-world-like)
  // -----------------------------------------------------------------------
  test("detects separators in a Spanish Outlook thread", () => {
    const html =
      "<p>Perfecto, queda confirmado.</p>" +
      "<p>-----Mensaje original-----</p>" +
      "<p>De: Carmen &lt;c@test.es&gt;<br>Enviado el: lunes, 24 marzo<br>Asunto: Proyecto</p>" +
      "<p>Estimados, me pongo en contacto...</p>" +
      "<p>" + "x".repeat(200) + "</p>" +
      "<p>-----Mensaje original-----</p>" +
      "<p>De: Alejandro &lt;a@test.es&gt;<br>Enviado el: lunes, 24 marzo<br>Asunto: Re: Proyecto</p>" +
      "<p>Buenos días, le confirmo mi disponibilidad...</p>";
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(2);
  });

  // -----------------------------------------------------------------------
  // French forward chain (real-world-like)
  // -----------------------------------------------------------------------
  test("detects separators in a French forward chain", () => {
    const html =
      "<p>FYI voir ci-dessous</p>" +
      "<p>-----Message d'origine-----</p>" +
      "<p>De : Maxime &lt;m@test.fr&gt;<br>Date : mercredi 26 mars<br>Objet : TR: Alerte</p>" +
      "<p>Notre bucket de backups est impacté.</p>" +
      "<p>" + "x".repeat(200) + "</p>" +
      "<p>-----Message d'origine-----</p>" +
      "<p>De : support@cloud.io<br>Date : mercredi 26 mars<br>Objet : Alerte</p>" +
      "<p>Sévérité: ÉLEVÉ</p>";
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(2);
  });

  // -----------------------------------------------------------------------
  // Mixed thread: Gmail ES attribution + Outlook ES header
  // -----------------------------------------------------------------------
  test("handles mixed Gmail/Outlook Spanish thread", () => {
    const html =
      "<p>Gracias por la propuesta.</p>" +
      "<p>El mié, 26 mar 2025, Gabriela &lt;g@test.mx&gt; escribió:</p>" +
      "<blockquote><p>Estimado Ing. Reyes...</p></blockquote>" +
      "<p>" + "x".repeat(200) + "</p>" +
      "<p>El 26 mar 2025, Marco &lt;m@test.mx&gt; escribió:</p>" +
      "<blockquote><p>Con gusto le respondo...</p></blockquote>" +
      "<p>" + "x".repeat(200) + "</p>" +
      "<p>El 26 mar 2025, Gabriela &lt;g@test.mx&gt; escribió:</p>" +
      "<blockquote><p>Voy a presentar la propuesta...</p></blockquote>";
    const seps = findReplySeparators(html);
    expect(seps.length).toBe(3);
  });

  // -----------------------------------------------------------------------
  // Edge case: no separators
  // -----------------------------------------------------------------------
  test("returns empty array for email with no replies", () => {
    const html = "<p>Just a simple email with no quoted content at all.</p>";
    expect(findReplySeparators(html)).toEqual([]);
  });
});
