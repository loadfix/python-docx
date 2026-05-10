# Security — python-docx (loadfix fork)

This fork is a file-format library. It writes and reads `.docx` bytes.
It does **not** render, sandbox, or validate the *semantic* content
that Word will subsequently process. Callers who accept untrusted
input are responsible for their own sanitisation.

## altChunk payloads (HTML, RTF, MHTML, plain text)

An `altChunk` is a Word primitive that embeds a foreign payload in
the package and asks Word to substitute the payload's rendered content
for the `<w:altChunk>` marker on open. The substitution runs inside
Word's native import filters, not inside python-docx.

That means an attacker who controls the altChunk payload controls
whatever Word's filter chooses to interpret:

- **HTML / XHTML** (`add_html_chunk`) — Word's HTML import pipeline
  historically evaluates embedded scripts, external image URLs, and
  conditional-comment VML. It has been the vector for CVE-2017-11826,
  CVE-2018-0802, and more recent template-injection issues.
- **RTF** (`add_rtf_chunk`) — RTF's control-word grammar lets payloads
  carry embedded OLE objects, external data links, and remote
  templates. CVE-2017-0199 and CVE-2023-21716 are well-known RCE
  examples that were triggered simply by opening a document.
- **MHTML** (`add_mhtml_chunk`) — multi-part archives combine HTML
  plus related resources; the HTML portion has the same exposure as
  HTML altChunks.
- **Plain text** (`add_text_chunk`) — lowest-risk, but note Word will
  still interpret hyperlink-looking strings on autoformat.

### python-docx's stance

python-docx **does not** auto-sanitise altChunk payloads. None of the
helpers below inspect, rewrite, or strip the input:

- `Document.add_alt_chunk(content, content_type=..., match_src=...)`
- `Document.add_html_chunk(html, match_src=...)`
- `Document.add_text_chunk(text, encoding=..., match_src=...)`
- `Document.add_rtf_chunk(rtf, match_src=...)`
- `Document.add_mhtml_chunk(mhtml, match_src=...)`

This is deliberate — sanitisation is a content-policy concern that
belongs to the embedding application, not to the serializer.

### What callers should do

If the payload originated outside your trust boundary:

1. **Sanitise HTML** with a library like
   [`bleach`](https://github.com/mozilla/bleach) or
   [`nh3`](https://github.com/messense/nh3) with an explicit allow
   list of tags/attributes. Strip `<script>`, event handlers
   (`on*=`), `javascript:` / `data:` URIs, and external references.
2. **Reject RTF** unless you can cryptographically verify its origin.
   RTF has no practical sanitiser; there is no safe subset for
   opaque third-party input.
3. **Reject MHTML** by default — it multiplexes HTML plus arbitrary
   MIME parts. Unpack it, sanitise the HTML portion, drop the
   attachments, and re-emit a plain HTML altChunk if you must.
4. **Plain text is safe to embed** but will be rendered verbatim;
   consider HTML-escaping if the caller expects literal characters
   like `<` or `&` to round-trip.

### See also

- ECMA-376 Part 1 §17.17 (Glossary Document and Alternate-Format
  Import Parts)
- Microsoft's [Protected View](https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653)
  guidance — this mitigates opening altChunks from untrusted origins,
  but does not eliminate the risk.

## Reporting a vulnerability in python-docx itself

If you find a security issue in the loadfix fork (parser
memory-exhaustion, ZIP-path traversal, XML external-entity handling,
etc.), file an issue on the loadfix monorepo and mark it security.
Do **not** open a public PR that includes a working exploit; drop a
note that describes the class of issue and wait for a maintainer to
reach out.
