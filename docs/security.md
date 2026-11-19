# Security

PPTXjs turns untrusted files into HTML that runs in your page, so a malicious presentation must not be
able to run script, exfiltrate data, tamper with the host page or exhaust the browser. This document lists
the threats considered, what v1 got wrong, and how v2 handles each one. Every item has a regression test
(`test/integration/security.test.ts`, `test/unit/*`).

## Threat model

- **Attacker:** anyone who can get a `.pptx` displayed by an application that uses PPTXjs (uploads, links, shared drives).
- **Assets:** the host page's origin (cookies, tokens, DOM), the user's browser (CPU and memory), the user's privacy.
- **Out of scope:** vulnerabilities in the browser itself; deliberately disabling the protections described here.

## Findings and fixes

| # | Issue in v1 | Impact | v2 mitigation |
|---|---|---|---|
| S1 | XML text and attribute values were concatenated into HTML mostly unescaped. Font names, shape names, tooltips, alt text and bullet characters went into quoted attributes, and CDATA content was inserted raw. | Stored XSS: a crafted file runs JavaScript in the host origin | The XML parser decodes and re-encodes every text node and attribute value (`& < > " '` become entities). Document content can never break out of a text node or a quoted attribute, whatever code path uses it. CDATA is escaped too. |
| S2 | Hyperlink targets were only HTML-escaped (issue #3 fix). `javascript:` URLs still worked. | XSS on click | `sanitizeUrl()` allows only `http(s)`, `mailto`, `tel`, `ftp`, in-page anchors and relative URLs, after decoding entities and removing control and whitespace characters. Links between slides become in-viewer navigation. |
| S3 | External links used `target=_blank` without `rel`. | Reverse tabnabbing | `rel="noopener noreferrer"` |
| S4 | Linked videos were embedded as an unsandboxed `<iframe>` with any URL from the file. | Third-party content in the page: phishing, tracking, autoplay | Off by default (poster plus link). `externalMedia: true` embeds only `http(s)` URLs in a sandboxed iframe with `referrerpolicy="no-referrer"`. |
| S5 | `Object.prototype.set` was assigned at runtime (issues #22, #28, #30). | Prototype pollution: broke jQuery, Vue and ExtJS, and could change the behaviour of unrelated code | Removed. The parser drops elements and attributes named `__proto__`, `constructor` or `prototype`. `setTextByPathList` refuses those keys. |
| S6 | No limits on archive size. | Zip bomb: a few KB that expand to gigabytes crash the tab | `maxUncompressedSize` (512 MiB) and `maxEntries` (20 000), checked before inflating. Configurable through `limits`. |
| S7 | `getBase64ImageDimensions` busy-waited in a `do … while` loop for an image to load. | Main-thread freeze (DoS) | Image size is read from the file header synchronously. |
| S8 | A global `keydown` handler called `preventDefault()` on every key. | Broke typing and accessibility in the host page | Keyboard handling is scoped to the viewer; inputs are ignored; `keyboard: "global"` is opt-in. |
| S9 | `jsZipV2` option: fetched a script from a configurable URL with `$.getScript` and reloaded the page; `slideType: "revealjs"` loaded `./revealjs/reveal.js` at runtime. | Script injection if options come from user input; surprising reloads | Removed. v2 never loads scripts at runtime. |
| S10 | Video/audio object URLs were never revoked. | Memory leak | Tracked and revoked on reload and on `destroy()`. |
| S11 | Generated CSS classes were global. Two viewers, or the host page, could clash. | Style injection between documents on the same page | Generated CSS is scoped to the viewer instance. The slideshow `background` option accepts only plain colour values. |
| S12 | Relationship targets were resolved with `replace("../", "ppt/")`. | Wrong files read, crashes | Proper OPC resolution relative to the source part. External targets are never read from the package. |

## Recommendations for applications

- Serve the page with a Content Security Policy that forbids inline scripts (`script-src 'self'`). PPTXjs
  never needs inline script. It needs `img-src data: blob:` and `media-src blob:` for embedded pictures and
  media. Slides are positioned with inline `style` attributes and the viewer injects a scoped `<style>`
  element, so `style-src 'unsafe-inline'` is required. Inline styles cannot run code.
- Keep `externalMedia` off for files from untrusted sources.
- Lower `limits` if you only expect small files.
- Rendering in a Web Worker (`renderToHtml`) isolates parsing from the page's main thread.

## Reporting

See [SECURITY.md](../SECURITY.md).
