const {
  escapeHtml,
  sanitizeSelectionHtml,
  toHtmlFromText,
} = require("../src/shared/office-helpers");

// ---------------------------------------------------------------------------
// escapeHtml
// ---------------------------------------------------------------------------
describe("escapeHtml", () => {
  test("escapes all five special characters", () => {
    expect(escapeHtml('&<>"\''))
      .toBe("&amp;&lt;&gt;&quot;&#39;");
  });

  test("leaves normal text unchanged", () => {
    expect(escapeHtml("Hello world 123")).toBe("Hello world 123");
  });

  test("coerces non-string values to string", () => {
    expect(escapeHtml(42)).toBe("42");
    expect(escapeHtml(null)).toBe("null");
    expect(escapeHtml(undefined)).toBe("undefined");
  });
});

// ---------------------------------------------------------------------------
// toHtmlFromText
// ---------------------------------------------------------------------------
describe("toHtmlFromText", () => {
  test("wraps text in a pre-wrap div and escapes HTML", () => {
    expect(toHtmlFromText("Hello <world>")).toBe(
      '<div style="white-space: pre-wrap;">Hello &lt;world&gt;</div>'
    );
  });

  test("handles empty / null input", () => {
    expect(toHtmlFromText("")).toBe('<div style="white-space: pre-wrap;"></div>');
    expect(toHtmlFromText(null)).toBe('<div style="white-space: pre-wrap;"></div>');
  });
});

// ---------------------------------------------------------------------------
// sanitizeSelectionHtml — script / comment removal (original behaviour)
// ---------------------------------------------------------------------------
describe("sanitizeSelectionHtml", () => {
  test("removes <script> tags and their content", () => {
    const input = '<p>Hello</p><script>alert("xss")</script><p>World</p>';
    expect(sanitizeSelectionHtml(input)).toBe("<p>Hello</p><p>World</p>");
  });

  test("removes HTML comments", () => {
    const input = "<p>A</p><!-- secret --><p>B</p>";
    expect(sanitizeSelectionHtml(input)).toBe("<p>A</p><p>B</p>");
  });

  // -----------------------------------------------------------------------
  // Dangerous tags (iframe, embed, object, applet, form)
  // -----------------------------------------------------------------------
  test("removes <iframe> with content", () => {
    const input = '<p>OK</p><iframe src="evil.html">inside</iframe><p>OK</p>';
    expect(sanitizeSelectionHtml(input)).toBe("<p>OK</p><p>OK</p>");
  });

  test("removes self-closing <embed>", () => {
    const input = '<p>OK</p><embed src="evil.swf"/><p>OK</p>';
    expect(sanitizeSelectionHtml(input)).toBe("<p>OK</p><p>OK</p>");
  });

  test("removes <object> with content", () => {
    const input = '<p>OK</p><object data="x"><param name="a" value="b"></object><p>OK</p>';
    expect(sanitizeSelectionHtml(input)).toBe("<p>OK</p><p>OK</p>");
  });

  test("removes <applet> tags", () => {
    const input = "<applet code='Hack.class'>fallback</applet>";
    expect(sanitizeSelectionHtml(input)).toBe("");
  });

  test("removes <form> tags", () => {
    const input = '<form action="/steal"><input type="text"></form>';
    expect(sanitizeSelectionHtml(input)).toBe("");
  });

  // -----------------------------------------------------------------------
  // Event handler attributes
  // -----------------------------------------------------------------------
  test("removes onerror handler (double quotes)", () => {
    const input = '<img src="x" onerror="alert(1)">';
    const result = sanitizeSelectionHtml(input);
    expect(result).not.toContain("onerror");
    expect(result).toContain("<img");
  });

  test("removes onload handler (single quotes)", () => {
    const input = "<svg onload='alert(1)'>";
    const result = sanitizeSelectionHtml(input);
    expect(result).not.toContain("onload");
  });

  test("removes onclick handler (no quotes)", () => {
    const input = '<a href="#" onclick=alert(1)>click</a>';
    const result = sanitizeSelectionHtml(input);
    expect(result).not.toContain("onclick");
    expect(result).toContain("<a");
  });

  test("removes onmouseover handler", () => {
    const input = '<div onmouseover="fetch(\'evil\')">hover me</div>';
    const result = sanitizeSelectionHtml(input);
    expect(result).not.toContain("onmouseover");
    expect(result).toContain("<div");
  });

  // -----------------------------------------------------------------------
  // javascript: and data: URIs
  // -----------------------------------------------------------------------
  test("neutralises javascript: in href", () => {
    const input = '<a href="javascript:alert(1)">click</a>';
    const result = sanitizeSelectionHtml(input);
    expect(result).not.toContain("javascript:");
  });

  test("neutralises javascript: in src", () => {
    const input = '<img src="javascript:alert(1)">';
    const result = sanitizeSelectionHtml(input);
    expect(result).not.toContain("javascript:");
  });

  test("neutralises data:text/html in src", () => {
    const input = '<iframe src="data:text/html,<script>alert(1)</script>">';
    const result = sanitizeSelectionHtml(input);
    expect(result).not.toContain("data:text/html");
  });

  // -----------------------------------------------------------------------
  // Preservation of legitimate Outlook HTML
  // -----------------------------------------------------------------------
  test("preserves normal <img> tags without handlers", () => {
    const input = '<img src="cid:image001.png" width="200" height="100">';
    expect(sanitizeSelectionHtml(input)).toBe(input);
  });

  test("preserves Outlook mso-* styles", () => {
    const input = '<p style="mso-line-height-rule:exactly;">Text</p>';
    expect(sanitizeSelectionHtml(input)).toBe(input);
  });

  test("preserves <o:p> namespace tags", () => {
    const input = "<o:p>&#160;</o:p>";
    expect(sanitizeSelectionHtml(input)).toBe(input);
  });

  test("preserves tables and data-* attributes", () => {
    const input = '<table data-custom="val"><tr><td>Cell</td></tr></table>';
    expect(sanitizeSelectionHtml(input)).toBe(input);
  });

  test("preserves normal <a> links", () => {
    const input = '<a href="https://example.com">link</a>';
    expect(sanitizeSelectionHtml(input)).toBe(input);
  });
});
