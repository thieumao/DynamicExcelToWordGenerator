# Word processing design

Libraries (versions verified on PyPI on 2026-11-04): **docxtpl 0.20.2** (LGPL-2.1, Jinja2 templates
inside .docx), **python-docx 1.2.0** (MIT), **docxcompose 2.2.0** (MIT, merging), **lxml**.
No Microsoft Word installation is needed for generation.

## 1. Template authoring (placeholders)

Three placeholder styles are supported per template (`templates[].marker.style`):

| Style | Example in Word | Notes |
|---|---|---|
| `jinja` (recommended) | `{{ employee.full_name }}`, `{%tr for p in employee.projects %}`, `{%p if employee.vip %}` | Full power: loops in table rows (`{%tr %}`), paragraphs (`{%p %}`), cells (`{%tc %}`), runs (`{%r %}`), filters (`{{ x | format_date }}`). |
| `custom` | `<<Họ tên>>`, `[MaNV]`, `${code}` (delimiters configurable) | Converted on an internal copy to `{{ b["Họ tên"] }}`; works even when Word split the marker across runs with different formatting (the first run's formatting is kept). Can be mixed with Jinja tags in the same file. |
| `mergefield` | Word mail-merge fields `«Họ_tên»` (`MERGEFIELD`) | Both `w:fldSimple` and complex fields (`fldChar`/`instrText`) are converted, keeping the result run's formatting. |

Custom/MERGEFIELD names are mapped to data through `templates[].bindings`
(`{"Họ tên": employee.full_name}`) or automatically when the name equals a field path or uniquely
matches the last segment of one (case/accent/underscore-insensitive) – auto-bindings are reported as
information so the user can review them on the Templates page.

**Not supported (documented, reported when detected):** MERGEFIELD regions (`TableStart:`/`TableEnd:`
– use `{%tr for %}` instead), MERGEFIELDs nested inside `IF`/`NEXT` fields, placeholders inside
text boxes' alternate content may only be partially detected, placeholders inside charts/SmartArt.

### Forms supplied as-is (input spots)

Customers often hand over an **existing form** instead of a prepared template. Many of its inputs have
no marker at all. The user picks these files in the app (Templates page *Import*, or *Add Word
files…* on the Generate page). `infrastructure/word/spots.py` then detects three kinds of
**input spots** on the internal compiled copy; the customer's file is never modified:

| Kind | Detected from | Name used for binding | How it is filled |
|---|---|---|---|
| Content control (`content_control`) | `w:sdt` with text, rich text, date, drop-down/combo, check box or picture content | the control's *tag*, else its *title*, else `control_N` | The control stays a form control. `showingPlcHdr` and the placeholder style are removed. A custom-XML `w:dataBinding` is also removed, because otherwise Word would overwrite the value on open. A stale date `fullDate` is removed too. The first text run receives the value, and the other runs are emptied. Check boxes get `☒/☐` and `w14:checked`. A **picture** control gets the image of an `image` field: the frame's placeholder picture is replaced and the image is fitted into the frame size (`wp:extent`), keeping its aspect ratio, unless the field sets `width_mm`/`height_mm`. No image in Excel leaves the frame empty. |
| Legacy form field (`form_field`) | `FORMTEXT` / `FORMCHECKBOX` / `FORMDROPDOWN` fields (`w:ffData/w:name`) | the field's bookmark name | The field result text is replaced. For a check box, `w:checked w:val` is set. For a **drop-down**, `w:ddList/w:result` selects the entry equal to the value: an exact match first, then ignoring case and spaces, then ignoring accents, and only when that match is unique. An empty value keeps the form's own selection. A value that is not one of the entries stops that document with `render.value_error`; it is never guessed. Use a `map_values` transform to translate codes into entries. |
| Blank (`blank`) | A run of `.`, `_` or `…` (minimum `blank_min`, where `…` counts as 3), or underlined spaces, after a label such as `Họ và tên: ........`; also an empty table cell to the right of a label cell | the label text, e.g. `Họ và tên`; a repeated label gets `tháng (2)` | The filler is replaced by the value. Text before and after it, and its run formatting, are kept. |

Rules:

* **Detection never changes anything.** A spot is converted only when it is bound:
  * explicitly, through `templates[].bindings` (`{"Họ và tên": nhan_vien.ho_ten}`);
  * or, for controls and form fields only, automatically by name. This uses the same rule as
    custom markers.

  Blanks always need an explicit binding. This keeps signature lines, "for office use only" boxes
  and date lines that must stay empty exactly as they are. Unbound spots are reported as
  information (`template.spot_unbound`, `template.blanks_unbound`).
* Detection can be switched off per template: `marker.content_controls`, `marker.form_fields`,
  `marker.blanks` and `marker.blank_min`. It works alongside any marker style, so one form can mix
  `{{ }}`/`<<…>>` markers with blanks and controls.
* Check boxes accept `true/false`, `1/0`, `x`, `có/không`, `yes/no` and similar values. In a template
  they are available as the filters `checkbox_symbol` / `checkbox_value`.
* The Templates page lists every spot with:
  * its kind;
  * its status ("will be filled" / "left as is");
  * a field combo box;
  * the surrounding sentence, so the user can tell e.g. which `tháng` is meant.

  Changing a binding rescans the template.

**Limitations of spot detection:**
* Drop-down values are checked when each document is filled, not in the *Validate* pass, so a bad
  value is reported for that record only.
* A picture control bound to a field that is not of type `image` shows the text value and is
  reported (`template.picture_not_image`).
* Text inside fields or inside other controls is never treated as a blank.
* Only the first 300 empty label/value table cells per part are considered. List-style tables
  (header row + empty rows) are not detected as blanks; use `{%tr for %}` for those.
* `.doc` (Word 97–2003) files must first be saved as `.docx` in Word.

## 2. Inspection (before generation)

`infrastructure/word/inspector.py` reads `document.xml`, headers, footers, footnotes and endnotes,
applies docxtpl's own XML patching (so detection matches rendering exactly, including split runs
and `{%tr %}` handling), parses the result with Jinja and walks the AST:

* `{{ a.b.c }}` → value reference `a.b.c`; `{% if %}` → condition; `{% for p in a.list %}` → loop,
  and `p.name` inside the loop becomes `a.list[].name`; `loop.*` and functions are ignored.
* Every reference is checked against the **context schema** derived from the profile (fields,
  collections with item fields, aggregations, `report.*`, `unit.*`, group paths for grouped units).
  Unknown paths and unbound markers are reported with "did you mean" suggestions **before** any
  document is generated (policy `missing_placeholder: error | warn | empty`).
* Template syntax errors (unclosed `{% if %}`, a `{%p %}` that is not alone in its paragraph…) are
  reported with the surrounding document text.
* Non-Jinja markers left in a Jinja template (`«Name»`, `<<Name>>`) produce a warning suggesting the
  right marker style.

## 3. Rendering

`TemplateRenderer` (one per worker) renders each template with **its own context**:

* the unit context (record, group or dataset) + `report` + `unit`;
* `context_root` (optional) exposes a sub-object at the top level (`{{ name }}` instead of
  `{{ employee.name }}`);
* `b` = resolved bindings for markers/MERGEFIELDs;
* `include_if` decides whether the template is part of the unit's output at all.

Rendering uses an immutable sandboxed Jinja environment with **autoescape** (values containing `<`,
`>` or `&` cannot corrupt the XML), all transforms as filters, a `finalize` hook that formats
`date` → `dd/mm/yyyy`, `Decimal` → locale decimal separator, booleans → Có/Không, `None` → empty.
Text values are NFC-normalised and stripped of characters invalid in XML; `\n` becomes a line
break. `ImageRef` values (fields with `type: image`) become inline pictures when the path is inside
the allowed folder, the type is png/jpg/gif/bmp and the size is under the limit; otherwise a warning
is reported and the picture is omitted. A tracking `Undefined` records placeholders that had no
value at run time; with policy `error` the document is not saved.

Rendering touches only text runs and Jinja-tagged structures: fonts, styles, tables, images,
headers/footers, section properties, numbering and hyperlinks of the template are preserved.
`tests/regression/test_formatting_preservation.py` compares fonts, run formatting, styles, tables,
merged cells, numbering, inline images, header images, page breaks, section orientation/margins,
footer fields and hyperlinks between template and output.

## 4. Merging – evaluation and design

### docxcompose behaviour (read from the 2.2.0 source)

`Composer.append()` copies body content, styles, numbering (restarting the first list of each
appended document), images, shapes, diagrams, footnotes and referenced parts. By design it:

1. **drops the body-level `w:sectPr` of each appended document** – its last section's page size,
   orientation and margins are lost;
2. **removes header/footer references** of appended documents – the first document's
   headers/footers apply to everything;
3. maps style IDs to the first document's definitions unless `preserve_styles=True`.

docxtpl's `new_subdoc()` uses the same machinery, so it has the same limitations.

### DocGen's merge engine (`infrastructure/word/merger.py`)

Per template (`templates[].section`):

| Option | Behaviour |
|---|---|
| `mode: preserve` (default) | Before appending, the master's current section properties are moved into the last paragraph (closing that section). Appended content is inserted by docxcompose, but header/footer references are **re-targeted to copies** of the source header/footer parts (with their images and links) instead of being removed. The appended document's own body `sectPr` becomes the new final section. Result: every template keeps its orientation, margins, page size, headers and footers. |
| `inherit_missing_headers: false` (default) | Header/footer types the part does not define get an explicit **empty** header/footer, so a part without a footer does not show the previous part's footer (Word's default inheritance). `true` restores Word's inheritance. |
| `mode: inherit` | Plain docxcompose: the first template's layout and headers/footers apply. A warning is emitted if the part had a different page size/orientation. |
| `break_before: page / section / none` | Page break paragraph, or a section break (`nextPage`, `oddPage`, `evenPage`, `continuous` via `start`). |
| `restart_page_numbering` | Sets `w:pgNumType w:start="1"` for that part's section; otherwise numbering continues. |
| `merge.preserve_styles` | Passes docxcompose's `preserve_styles` (keep appended documents' style definitions). |
| `merge.update_fields_on_open` | Sets `w:updateFields` when TOC/PAGEREF/NUMPAGES fields exist so Word refreshes them on open. |

### Proof-of-concept results (`tests/poc/test_merge_poc.py`)

Three documents (portrait cover with header + PAGE field footer; landscape details with 12 mm
margins, header image, footer, table, numbered list, picture, hyperlink; portrait appendix with a
numbered list and a custom style) merged in `preserve` mode:

| Check | Result |
|---|---|
| 3 sections, orientations portrait/landscape/portrait, margins of part 2 kept | pass |
| Each section shows its own header text; header image resolves to an image part | pass |
| Part without footer gets an empty footer (no inheritance) | pass |
| Table, numbered list items (restarted per document), body image, hyperlink, custom style | pass |
| ZIP/relationship verification and python-docx reopen | pass |
| **Independent layout check:** LibreOffice 24.2 renders exactly 3 pages, each page shows its own header, numbering restarts (`1. Bước một`, `1. Mục A`), part 3 has no footer | pass |

The same test file covers `inherit` mode (single section, first header, orientation warning) and
section breaks with restarted numbering.

### Office conformance (Microsoft Open XML SDK)

`tools/ooxml-validator` wraps `DocumentFormat.OpenXml` 3.4.1's `OpenXmlValidator` (schema and
semantic rules Microsoft uses for Office documents). Results on 2026-11-04 for 42 files (all
templates, every output of customers A and B in `separate` + `merged` mode, the rich regression
template merged in `preserve`, `section` and `inherit` mode):

| Target | Errors |
|---|---|
| Office 2010 | 0 |
| Office 2019 | 0 |
| Microsoft 365 | 0 |
| Office 2007 | only `w:tblLook` 2010+ attributes coming from python-docx's default template (not produced by DocGen; Word 2007 is not a target) |

Outputs of the PyInstaller-built application are validated the same way in the packaging smoke test.

### Templates saved by other editors (finding and fix)

Re-saving the sample templates in LibreOffice (docx → odt → docx) showed that unbound MERGEFIELDs are
**flattened to plain text `«Name»`**. Before the fix these markers were silently copied into the
outputs. Now, in `mergefield` mode, both real fields and their visible `«Name»` form are converted,
templates without any placeholder produce a warning, and output verification checks style-specific
leftovers (`«…»` and `MERGEFIELD` instructions for mergefield templates, the configured delimiters
for custom markers) in addition to `{{ }}`/`{% %}`. Covered by
`tests/regression/test_libreoffice_templates.py`.

### Known limitations (not hidden)

* **Word itself was not available in this environment.** Structural checks and LibreOffice
  rendering pass, but final acceptance must include opening outputs in Microsoft Word on Windows
  (checklist in [TEST_STRATEGY.md](TEST_STRATEGY.md)).
* `w:evenAndOddHeaders` is a document-wide setting: the first template's setting applies; a
  warning is emitted when parts differ.
* Tables of contents and cross-references are not recalculated by DocGen; Word updates them when
  opening the file (`updateFields`), LibreOffice may need a manual update.
* Page numbering across parts continues by default; per-part restart is opt-in.
* Style conflicts: with default settings the first template's style definitions win for identical
  style IDs. Use `preserve_styles` or distinct style names when templates must look different.
* Content controls are filled as described in §1 and checked by the Open XML SDK. Controls that are
  not bound in DocGen keep everything, including any custom XML binding. Comments and tracked
  changes are copied as-is but not specifically tested.
* Text boxes/shapes and SmartArt are copied by docxcompose; placeholders inside them are rendered by
  docxtpl when they are in the main document part, but are not part of the regression suite.

### Alternatives if a customer feature cannot be preserved

1. **Skeleton template** – one master .docx with `{{p subdoc }}` placeholders for docxtpl subdocs
   (same engine; useful when a fixed outer layout is required).
2. **Separate documents** (`output: separate` or `both`) – each template keeps 100 % of its own
   formatting; merging can then be done by the customer in Word.
3. **`altChunk`** – Word merges on open; rejected as default because it requires Word to render and
   is not supported by LibreOffice or most viewers.
4. **Optional Word automation (COM)** – would give pixel-perfect merging but needs Word and Windows;
   intentionally *not* part of the core. If ever added it must stay an isolated optional plug-in.

## 5. Output verification

Every generated or merged file is checked before it is moved into place: ZIP CRC test, required
parts, every internal relationship target exists, duplicate drawing ids (warning), **no leftover
`{{ }}`/`{% %}` markers** in body/headers/footers/footnotes, and python-docx can reopen it. A failure
marks the unit as failed and the file is not delivered.
