"""Word merge engine built on docxcompose.

docxcompose (2.2.0) appends body content, styles, numbering, images, footnotes
and other referenced parts, but by design it:

* drops the body-level ``w:sectPr`` of each appended document (orientation,
  margins and page size of the appended document's last section are lost), and
* removes header/footer references of appended documents (the first document's
  headers/footers apply everywhere).

This module keeps docxcompose's behaviour as the ``inherit`` mode and adds a
``preserve`` mode that keeps every appended part's section properties and its
own headers/footers:

1. the master's current body ``sectPr`` is copied into the last paragraph,
   which closes the master's section with its own properties;
2. the appended content is inserted by docxcompose, with header/footer
   references *re-targeted* to copies of the source header/footer parts
   (instead of removed);
3. the appended document's body ``sectPr`` (with re-targeted references)
   becomes the new final section.

Known limitations are listed in docs/WORD_PROCESSING_DESIGN.md and are also
returned as warnings in :class:`MergeReport`.
"""

from __future__ import annotations

from collections.abc import Sequence
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

from docx import Document
from docx.document import Document as DocxDocument
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.parts.hdrftr import FooterPart, HeaderPart
from docxcompose.composer import Composer
from docxcompose.image import ImageWrapper
from lxml import etree

SectionMode = Literal["inherit", "preserve"]
BreakMode = Literal["none", "page", "section"]
SectionStart = Literal["nextPage", "continuous", "oddPage", "evenPage"]

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(frozen=True, slots=True)
class MergePart:
    """One rendered document to be merged, with its layout options."""

    path: Path
    label: str = ""
    section_mode: SectionMode = "inherit"
    break_before: BreakMode = "page"
    section_start: SectionStart = "nextPage"
    restart_page_numbering: bool = False
    inherit_missing_headers: bool = False
    """preserve mode only: True keeps Word's default (missing header/footer types are inherited
    from the previous part); False gives such types an explicit empty header/footer."""


@dataclass(slots=True)
class MergeReport:
    output: Path
    parts: int
    sections: int
    warnings: list[str] = field(default_factory=list)


class _SectionAwareComposer(Composer):  # type: ignore[misc]
    """docxcompose Composer that can re-target headers/footers instead of dropping them."""

    def __init__(self, doc: DocxDocument, preserve_styles: bool = False) -> None:
        super().__init__(doc, preserve_styles=preserve_styles)
        self._preserve_current = False
        self._blank_parts: dict[str, str] = {}

    # -- overrides -------------------------------------------------------
    def remove_header_and_footer_references(self, doc: DocxDocument, element: Any) -> None:
        if not self._preserve_current:
            super().remove_header_and_footer_references(doc, element)
            return
        refs = element.xpath(".//w:headerReference|.//w:footerReference")
        for ref in refs:
            ref.set(qn("r:id"), self._copy_hdrftr(doc, ref.get(qn("r:id"))))

    def fix_header_and_footers(self, doc: DocxDocument) -> None:
        if not self._preserve_current:
            super().fix_header_and_footers(doc)

    def fix_section_types(self, doc: DocxDocument) -> None:
        if not self._preserve_current:
            super().fix_section_types(doc)

    # -- helpers ---------------------------------------------------------
    def _copy_hdrftr(self, src_doc: DocxDocument, src_rid: str) -> str:
        """Copy a header/footer part (and its images/links) into the master; return new rId."""
        src_rel = src_doc.part.rels[src_rid]
        src_part = src_rel.target_part
        is_header = src_rel.reltype == RT.HEADER
        pkg = self.doc.part.package
        template = "/word/header%d.xml" if is_header else "/word/footer%d.xml"
        partname = pkg.next_partname(template)
        element = parse_xml(src_part.blob)
        part_cls = HeaderPart if is_header else FooterPart
        content_type = CT.WML_HEADER if is_header else CT.WML_FOOTER
        new_part = part_cls(partname, content_type, element, pkg)

        rid_map: dict[str, str] = {}
        for rel in list(src_part.rels.values()):
            if rel.is_external:
                rid_map[rel.rId] = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.IMAGE:
                img_part = rel.target_part
                existing = pkg.image_parts._get_by_sha1(img_part.sha1)
                if existing is None:
                    existing = pkg.image_parts._add_image_part(ImageWrapper(img_part))
                rid_map[rel.rId] = new_part.relate_to(existing, RT.IMAGE)
            else:
                new_rel = self.add_relationship(src_part, new_part, rel)
                rid_map[rel.rId] = new_rel.rId

        for el in element.iter():
            for attr, value in list(el.attrib.items()):
                if str(attr).startswith(f"{{{R_NS}}}") and str(value) in rid_map:
                    el.set(attr, rid_map[str(value)])
        # styles referenced inside header/footer must exist in the master
        self.add_styles(src_doc, element)
        result: str = self.doc.part.relate_to(new_part, src_rel.reltype)
        return result

    def append_document(
        self,
        doc: DocxDocument,
        *,
        section_mode: SectionMode,
        break_before: BreakMode,
        section_start: SectionStart,
        restart_page_numbering: bool,
        inherit_missing_headers: bool = False,
    ) -> None:
        body = self.doc.element.body
        if section_mode == "preserve":
            self._close_current_section()
            sections_before = len(self._all_sectprs()) - 1  # minus the body sectPr
            self._preserve_current = True
            try:
                self.append(doc)
            finally:
                self._preserve_current = False
            src_sectpr = doc.element.body.find(qn("w:sectPr"))
            if src_sectpr is not None:
                new_sectpr = deepcopy(src_sectpr)
                for ref in new_sectpr.xpath("./w:headerReference|./w:footerReference"):
                    ref.set(qn("r:id"), self._copy_hdrftr(doc, ref.get(qn("r:id"))))
                old = body.find(qn("w:sectPr"))
                if old is not None:
                    body.replace(old, new_sectpr)
                else:
                    body.append(new_sectpr)
            self._set_final_section_start(section_start)
            self._set_page_numbering(restart_page_numbering)
            if not inherit_missing_headers:
                first_new = self._all_sectprs()[sections_before]
                self._fill_missing_hdrftr(first_new)
            return

        if break_before == "page":
            self._append_page_break()
        elif break_before == "section":
            self._close_current_section()
            self._set_final_section_start(section_start)
        self.append(doc)
        if restart_page_numbering and break_before == "section":
            self._set_page_numbering(True)

    def _all_sectprs(self) -> list[Any]:
        body = self.doc.element.body
        result: list[Any] = body.xpath("./w:p/w:pPr/w:sectPr")
        body_sectpr = body.find(qn("w:sectPr"))
        if body_sectpr is not None:
            result.append(body_sectpr)
        return result

    def _blank_hdrftr(self, kind: str) -> str:
        """rId of a shared empty header/footer part, created on first use."""
        if kind not in self._blank_parts:
            pkg = self.doc.part.package
            if kind == "header":
                blank: Any = HeaderPart.new(pkg)
                reltype = RT.HEADER
            else:
                blank = FooterPart.new(pkg)
                reltype = RT.FOOTER
            self._blank_parts[kind] = self.doc.part.relate_to(blank, reltype)
        return self._blank_parts[kind]

    def _fill_missing_hdrftr(self, sectpr: Any) -> None:
        """Give *sectpr* explicit empty headers/footers for types it does not define."""
        for kind in ("header", "footer"):
            tag = qn(f"w:{kind}Reference")
            present = {ref.get(qn("w:type")) for ref in sectpr.findall(tag)}
            for ref_type in ("default", "first", "even"):
                if ref_type in present:
                    continue
                ref = OxmlElement(f"w:{kind}Reference")
                ref.set(qn("w:type"), ref_type)
                ref.set(qn("r:id"), self._blank_hdrftr(kind))
                # header/footer references must be the first children of sectPr
                sectpr.insert(0, ref)

    def _last_block(self) -> Any:
        body = self.doc.element.body
        children = [c for c in body if c.tag != qn("w:sectPr")]
        return children[-1] if children else None

    def _close_current_section(self) -> None:
        """Copy the body sectPr into the last paragraph, ending the current section."""
        body = self.doc.element.body
        body_sectpr = body.find(qn("w:sectPr"))
        if body_sectpr is None:
            return
        last = self._last_block()
        if last is None or last.tag != qn("w:p") or last.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            new_p = OxmlElement("w:p")
            body_sectpr.addprevious(new_p)
            last = new_p
        ppr = last.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            last.insert(0, ppr)
        closing = deepcopy(body_sectpr)
        # sectPr must be the last child of pPr except for pPrChange
        change = ppr.find(qn("w:pPrChange"))
        if change is not None:
            change.addprevious(closing)
        else:
            ppr.append(closing)

    def _set_final_section_start(self, start: SectionStart) -> None:
        sectpr = self.doc.element.body.find(qn("w:sectPr"))
        if sectpr is None:
            return
        type_el = sectpr.find(qn("w:type"))
        if type_el is None:
            type_el = OxmlElement("w:type")
            # w:type must appear after header/footer refs and footnotePr/endnotePr
            anchor = None
            for tag in ("w:endnotePr", "w:footnotePr", "w:footerReference", "w:headerReference"):
                found = sectpr.findall(qn(tag))
                if found:
                    anchor = found[-1]
                    break
            if anchor is not None:
                anchor.addnext(type_el)
            else:
                sectpr.insert(0, type_el)
        type_el.set(qn("w:val"), start)

    def _set_page_numbering(self, restart: bool) -> None:
        sectpr = self.doc.element.body.find(qn("w:sectPr"))
        if sectpr is None:
            return
        pg = sectpr.find(qn("w:pgNumType"))
        if restart:
            if pg is None:
                pg = OxmlElement("w:pgNumType")
                # place before cols/titlePg etc.; appending is tolerated by Word but keep order
                cols = sectpr.find(qn("w:cols"))
                if cols is not None:
                    cols.addprevious(pg)
                else:
                    sectpr.append(pg)
            pg.set(qn("w:start"), "1")
        elif pg is not None and pg.get(qn("w:start")) is not None:
            del pg.attrib[qn("w:start")]

    def _append_page_break(self) -> None:
        body = self.doc.element.body
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        r.append(br)
        p.append(r)
        sectpr = body.find(qn("w:sectPr"))
        if sectpr is not None:
            sectpr.addprevious(p)
        else:
            body.append(p)


def _enable_update_fields(doc: DocxDocument) -> None:
    """Ask Word to refresh fields (TOC, page refs) when the merged file is opened."""
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


def _uses_even_odd_headers(doc: DocxDocument) -> bool:
    return doc.settings.element.find(qn("w:evenAndOddHeaders")) is not None


def _has_fields(doc: DocxDocument, names: Sequence[str]) -> bool:
    xml = etree.tostring(doc.element.body, encoding="unicode")
    return any(name in xml for name in names)


class WordMergeEngine:
    """Merge rendered DOCX files into one document."""

    def __init__(self, *, preserve_styles: bool = False, update_fields_on_open: bool = True) -> None:
        self.preserve_styles = preserve_styles
        self.update_fields_on_open = update_fields_on_open

    def merge(self, parts: Sequence[MergePart], output: Path) -> MergeReport:
        if not parts:
            raise ValueError("Nothing to merge")
        warnings: list[str] = []
        master = Document(str(parts[0].path))
        composer = _SectionAwareComposer(master, preserve_styles=self.preserve_styles)
        master_even_odd = _uses_even_odd_headers(master)
        needs_field_update = _has_fields(master, ("TOC", "PAGEREF", "NUMPAGES"))

        for part in parts[1:]:
            doc = Document(str(part.path))
            if _uses_even_odd_headers(doc) != master_even_odd:
                warnings.append(
                    f"'{part.label or part.path.name}': even/odd header setting differs from the "
                    "first document; the first document's setting applies to the merged file."
                )
            if part.section_mode == "inherit" and len(doc.sections) == 1:
                src = doc.sections[0]
                dst = master.sections[-1]
                if (src.orientation, src.page_width, src.page_height) != (
                    dst.orientation,
                    dst.page_width,
                    dst.page_height,
                ):
                    warnings.append(
                        f"'{part.label or part.path.name}': page size/orientation differs from the "
                        "previous part but section mode is 'inherit'; use section mode 'preserve' "
                        "to keep it."
                    )
            needs_field_update = needs_field_update or _has_fields(doc, ("TOC", "PAGEREF", "NUMPAGES"))
            composer.append_document(
                doc,
                section_mode=part.section_mode,
                break_before=part.break_before,
                section_start=part.section_start,
                restart_page_numbering=part.restart_page_numbering,
                inherit_missing_headers=part.inherit_missing_headers,
            )

        if self.update_fields_on_open and needs_field_update:
            _enable_update_fields(master)
        output.parent.mkdir(parents=True, exist_ok=True)
        composer.save(str(output))
        return MergeReport(output=output, parts=len(parts), sections=len(master.sections), warnings=warnings)
