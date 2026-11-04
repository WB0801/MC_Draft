# Editable Word export

**Goal:** Add Word download for the current, manually editable MC script. Multi-file importing remains out of scope.

**Design:** A new Download Word button beside the existing text download. Export the textarea exactly as it currently stands, never regenerate it. Produce a real editable DOCX entirely in the browser without adding a CDN dependency. Use a minimal uncompressed ZIP/OOXML writer in one separate module; no server or data upload. Black text, a centered bold first-line title, 14-point Chinese KaiTi with Times New Roman for Latin text, readable line spacing and explicit page margins. Keep competition introductions with the first recipient and avoid splitting individual paragraphs, without forcing an entire long list onto one page.

**Files:** index.html (button and download handler), word-export.js (pure export builder), tests/word-export.test.cjs and existing app tests (regressions), README.md (usage).

## Tasks
- [x] Write failing ZIP/XML, text fidelity and download integration tests.
- [x] Implement the exporter and download handler; keep current parsing/generation unchanged.
- [x] Generate representative short and multi-page DOCX with the real exporter; inspect extracted text and render every page.
- [ ] Run the full suite, syntax/whitespace checks and independent review; push a reviewable branch/PR and update the local test package.

## Review focus
- Manual edits, empty lines, XML special characters, tabs and emoji survive export.
- The download uses current textarea content even after settings have changed.
- Empty input and a missing export module fail with a clear message.
- Long lists can continue across pages; an introduction should not be stranded at a page bottom.
- The ZIP has valid CRCs, offsets and content types and opens in actual Word-compatible software.

## Verification and review
- 47 Node tests pass, including the existing 33 regressions. The browser classic-script global entry point was also exercised without Node globals.
- Actual output ZIP CRCs, all XML parts, exact paragraph text and Title style independently validated with zipfile, lxml and python-docx.
- Short one-page and long four-page samples rendered with LibreOffice; every page visually inspected for text, spacing and list continuation.
- Independent review found out-of-order OOXML paragraph properties (Important) and a repeated external representative introduction missing from the pagination heuristic (Minor). Both fixed with failing regression tests first; full suite green after the fixes. Final documents regenerated and rendered again.
- Word fonts depend on installed reader fonts; the intended styles explicitly request KaiTi for Chinese and Times New Roman for Latin. No font embedding or new CDN dependency added.
