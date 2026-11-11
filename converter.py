"""
converter.py - PowerPoint (.pptx) -> Word (.docx) conversion engine.

Pipeline (mirrors the architecture diagram):

    .pptx --> extraction (python-pptx) --> document building (python-docx) --> .docx

The engine has no UI code, so it can be tested on its own or reused from a CLI.

How slides are represented in Word
----------------------------------
* Slide title        -> Heading 1 ("Slide 3: Title" if numbering is on)
* Title slide        -> Word "Title" / "Subtitle" styles, then a page break
* Bullets            -> real Word bullet lists, with the original nesting level
* Numbered lists     -> numbered text that restarts per text box (no runaway numbering)
* Plain text boxes   -> normal paragraphs
* Bold/italic/underline and hyperlinks are preserved
* Tables             -> real Word tables (header row shaded, merged cells kept)
* Pictures           -> inline images, scaled to page width, cropping applied
* Charts             -> a data table of the chart values (Word cannot embed a PPT chart)
* Grouped shapes     -> flattened and processed recursively
* Speaker notes      -> optional shaded "Notes" block
* SmartArt/video/OLE -> a clear placeholder line so nothing disappears silently
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.run import Run as DocxRun
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.text.text import _Run as PptxRun

# --------------------------------------------------------------------------- #
# Options / result containers
# --------------------------------------------------------------------------- #

SEPARATOR_RULE = "rule"          # thin line between slides (compact)
SEPARATOR_PAGE = "page"          # every slide starts on a new page
SEPARATOR_NONE = "none"          # just spacing


@dataclass
class ConvertOptions:
    include_slide_numbers: bool = True      # "Slide 3: Title" instead of "Title"
    include_images: bool = True
    include_tables: bool = True
    include_charts: bool = True
    include_notes: bool = False             # speaker notes
    include_hidden_slides: bool = False
    slide_separator: str = SEPARATOR_RULE
    cover_from_first_slide: bool = True     # first slide -> Title/Subtitle styles
    page_numbers: bool = True               # "Page X" footer
    landscape: bool = False


@dataclass
class ConvertResult:
    output_path: Path
    slides_total: int = 0
    slides_converted: int = 0
    images: int = 0
    tables: int = 0
    charts: int = 0
    warnings: list[str] = field(default_factory=list)


ProgressFn = Callable[[int, int, str], None]

# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #

_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_ROW_TOLERANCE = 274320  # 0.3 inch in EMU: shapes this close vertically are "same row"
_DOCX_NATIVE_IMAGES = {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff"}

BODY_FONT = "Calibri"
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
MUTED = RGBColor(0x59, 0x59, 0x59)
LINK = RGBColor(0x05, 0x63, 0xC1)


def clean(text: str) -> str:
    """Remove characters that are illegal in XML (they crash python-docx)."""
    return _CONTROL_CHARS.sub("", text or "")


def _local(el) -> str:
    return etree.QName(el).localname


def _set_style_font(style, name: str, size: Optional[float] = None,
                    bold: Optional[bool] = None, color: Optional[RGBColor] = None):
    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)
    # Remove theme font attributes, which would override the names set above.
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color


def _shade(cell_or_par_element, fill: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell_or_par_element.append(shd)


def _paragraph_border(paragraph, side: str = "bottom", color: str = "BFBFBF",
                      size: int = 6, space: int = 1):
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    b = OxmlElement(f"w:{side}")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(size))
    b.set(qn("w:space"), str(space))
    b.set(qn("w:color"), color)
    borders.append(b)


def _add_field(paragraph, instruction: str):
    """Insert a Word field (e.g. PAGE) into a paragraph."""
    run = paragraph.add_run()
    for kind, text in (("begin", None), (None, instruction), ("separate", None), (None, "1"), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        elif text == instruction:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = f" {instruction} "
        else:
            el = OxmlElement("w:t")
            el.text = text
        run._r.append(el)
    return run


# --------------------------------------------------------------------------- #
# The converter
# --------------------------------------------------------------------------- #

class PptxToDocx:
    def __init__(self, options: Optional[ConvertOptions] = None):
        self.opt = options or ConvertOptions()
        self.doc: Document = None  # type: ignore[assignment]
        self.result: ConvertResult = None  # type: ignore[assignment]
        self.slide_w = 1
        self.usable_w_cm = 17.0
        self.usable_h_cm = 20.0
        self._slide_no = 0

    # ------------------------------------------------------------------ API
    def convert(self, src: str | Path, dst: str | Path,
                progress: Optional[ProgressFn] = None) -> ConvertResult:
        src, dst = Path(src), Path(dst)
        if src.suffix.lower() != ".pptx":
            raise ValueError(
                "Only .pptx files are supported. Open older .ppt files in PowerPoint "
                "and use File > Save As > PowerPoint Presentation (*.pptx)."
            )

        prs = Presentation(str(src))
        self.slide_w = int(prs.slide_width) or 1
        self.result = ConvertResult(output_path=dst, slides_total=len(prs.slides))
        self._setup_document(prs, src)

        total = len(prs.slides)
        first_written = True
        for idx, slide in enumerate(prs.slides, start=1):
            self._slide_no = idx
            if progress:
                progress(idx - 1, total, f"Slide {idx} of {total}")

            hidden = slide._element.get("show") == "0"
            if hidden and not self.opt.include_hidden_slides:
                continue

            try:
                if not first_written and self.opt.slide_separator == SEPARATOR_PAGE:
                    self.doc.add_page_break()
                self._convert_slide(slide, idx, is_first=first_written)
                self.result.slides_converted += 1
                first_written = False
            except Exception as exc:  # keep going: one bad slide must not kill the job
                self.result.warnings.append(f"Slide {idx}: could not be fully converted ({exc})")

        dst.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(dst))
        if progress:
            progress(total, total, "Done")
        return self.result

    # ------------------------------------------------------------ document
    def _setup_document(self, prs, src: Path):
        doc = Document()
        self.doc = doc

        sec = doc.sections[0]
        if self.opt.landscape:
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
        else:
            sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)  # A4
        for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, side, Cm(2.0))
        self.usable_w_cm = (sec.page_width - sec.left_margin - sec.right_margin) / 360000
        self.usable_h_cm = (sec.page_height - sec.top_margin - sec.bottom_margin) / 360000 - 2

        styles = doc.styles
        _set_style_font(styles["Normal"], BODY_FONT, 11)
        styles["Normal"].paragraph_format.space_after = Pt(6)
        styles["Normal"].paragraph_format.line_spacing = 1.1

        _set_style_font(styles["Heading 1"], BODY_FONT, 17, True, ACCENT)
        h1 = styles["Heading 1"].paragraph_format
        h1.space_before, h1.space_after, h1.keep_with_next = Pt(14), Pt(6), True

        _set_style_font(styles["Title"], BODY_FONT, 30, True, ACCENT)
        _set_style_font(styles["Subtitle"], BODY_FONT, 15, False, MUTED)
        for name in ("List Bullet", "List Bullet 2", "List Bullet 3"):
            _set_style_font(styles[name], BODY_FONT, 11)
            styles[name].paragraph_format.space_after = Pt(3)

        props = doc.core_properties
        try:
            props.title = prs.core_properties.title or src.stem
            props.author = prs.core_properties.author or ""
        except Exception:
            props.title = src.stem

        if self.opt.page_numbers:
            _set_style_font(styles["Footer"], BODY_FONT, 9, False, MUTED)
            fp = sec.footer.paragraphs[0]
            fp.style = styles["Footer"]
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            fp.add_run("Page ")
            _add_field(fp, "PAGE")

    # --------------------------------------------------------------- slide
    def _convert_slide(self, slide, number: int, is_first: bool):
        title_shape = self._find_title(slide)
        title_text = ""
        if title_shape is not None and title_shape.has_text_frame:
            title_text = re.sub(r"\s+", " ", clean(title_shape.text_frame.text.replace("\v", " "))).strip()

        is_cover = (
            is_first and number == 1 and self.opt.cover_from_first_slide
            and title_shape is not None
            and title_shape.is_placeholder
            and title_shape.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE
        )

        if is_cover:
            p = self.doc.add_paragraph(title_text or "Presentation", style="Title")
            p.paragraph_format.space_before = Pt(60)
        else:
            if self.opt.include_slide_numbers:
                heading = f"Slide {number}: {title_text}" if title_text else f"Slide {number}"
            else:
                heading = title_text or f"Slide {number}"
            self.doc.add_heading(heading, level=1)

        self._walk_shapes(slide.shapes, skip=title_shape, subtitle_style=is_cover)

        if self.opt.include_notes:
            self._add_notes(slide)

        if is_cover:
            self.doc.add_page_break()
        elif self.opt.slide_separator == SEPARATOR_RULE:
            sep = self.doc.add_paragraph()
            sep.paragraph_format.space_after = Pt(4)
            _paragraph_border(sep, "bottom")

    @staticmethod
    def _find_title(slide):
        try:
            if slide.shapes.title is not None:
                return slide.shapes.title
        except Exception:
            pass
        return None

    # -------------------------------------------------------------- shapes
    @staticmethod
    def _sort_key(shape):
        top = shape.top if shape.top is not None else 0
        left = shape.left if shape.left is not None else 0
        return (int(top) // _ROW_TOLERANCE, int(left))

    def _walk_shapes(self, shapes, skip=None, subtitle_style: bool = False):
        for shape in sorted(shapes, key=self._sort_key):
            if skip is not None and shape.shape_id == skip.shape_id and shape.name == skip.name:
                continue
            try:
                self._handle_shape(shape, subtitle_style)
            except Exception as exc:
                self.result.warnings.append(
                    f"Slide {self._slide_no}: skipped '{shape.name}' ({exc})"
                )

    def _handle_shape(self, shape, subtitle_style: bool):
        try:
            stype = shape.shape_type
        except NotImplementedError:
            stype = None

        if stype == MSO_SHAPE_TYPE.GROUP:
            self._walk_shapes(shape.shapes, subtitle_style=subtitle_style)
            return

        # Pictures (including pictures sitting in a content placeholder)
        if stype == MSO_SHAPE_TYPE.PICTURE or self._is_picture_placeholder(shape):
            if self.opt.include_images:
                self._add_picture(shape)
            return

        if getattr(shape, "has_table", False) and shape.has_table:
            if self.opt.include_tables:
                self._add_table(shape.table)
            return

        if getattr(shape, "has_chart", False) and shape.has_chart:
            if self.opt.include_charts:
                self._add_chart(shape.chart)
            return

        if stype == MSO_SHAPE_TYPE.MEDIA:
            self._placeholder_line(f"[Media file omitted: {shape.name}]")
            return
        if stype in (MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT):
            self._placeholder_line(f"[Embedded object omitted: {shape.name}]")
            return
        if self._is_smartart(shape):
            self._placeholder_line("[SmartArt diagram - not convertible, please recreate or screenshot it]")
            return

        if shape.has_text_frame and shape.text_frame.text.strip():
            self._add_text_frame(shape, subtitle_style)

    @staticmethod
    def _is_picture_placeholder(shape) -> bool:
        if not getattr(shape, "is_placeholder", False):
            return False
        return hasattr(shape, "image") and shape.__class__.__name__ in ("PlaceholderPicture", "Picture")

    @staticmethod
    def _is_smartart(shape) -> bool:
        try:
            return bool(shape._element.xpath(".//*[local-name()='relIds']"))
        except Exception:
            return False

    def _placeholder_line(self, text: str):
        p = self.doc.add_paragraph()
        r = p.add_run(text)
        r.italic = True
        r.font.color.rgb = MUTED
        r.font.size = Pt(9.5)

    # ---------------------------------------------------------------- text
    @staticmethod
    def _bullet_kind(paragraph, shape, single_paragraph: bool) -> Optional[str]:
        """Return 'bullet', 'number' or None for a PowerPoint paragraph."""
        ppr = paragraph._p.pPr
        if ppr is not None:
            for child in ppr:
                name = _local(child)
                if name == "buNone":
                    return None
                if name == "buChar":
                    return "bullet"
                if name == "buAutoNum":
                    return "number"
        # Body/content placeholders inherit bullets from the slide layout.
        if getattr(shape, "is_placeholder", False) and not single_paragraph:
            if shape.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
                return "bullet"
        return None

    def _add_text_frame(self, shape, subtitle_style: bool):
        tf = shape.text_frame
        paragraphs = [p for p in tf.paragraphs if clean(p.text).strip()]
        single = len(paragraphs) == 1
        counters: dict[int, int] = {}

        is_subtitle = (
            getattr(shape, "is_placeholder", False)
            and shape.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE
        )

        for para in paragraphs:
            level = max(0, min(int(para.level or 0), 4))
            kind = self._bullet_kind(para, shape, single)

            if is_subtitle and subtitle_style:
                dp = self.doc.add_paragraph(style="Subtitle")
            elif kind == "bullet":
                style = ["List Bullet", "List Bullet 2", "List Bullet 3"][min(level, 2)]
                dp = self.doc.add_paragraph(style=style)
            elif kind == "number":
                for deeper in [k for k in counters if k > level]:
                    del counters[deeper]
                counters[level] = counters.get(level, 0) + 1
                dp = self.doc.add_paragraph()
                dp.paragraph_format.left_indent = Cm(0.63 * (level + 1))
                dp.paragraph_format.first_line_indent = Cm(-0.63)
                dp.paragraph_format.space_after = Pt(3)
                dp.add_run(f"{counters[level]}.\t")
                dp.paragraph_format.tab_stops.add_tab_stop(Cm(0.63 * (level + 1)))
            else:
                dp = self.doc.add_paragraph()
                if level:
                    dp.paragraph_format.left_indent = Cm(0.63 * level)
                counters.clear()

            self._write_runs(dp, para)

    def _write_runs(self, dest_par, src_par, size_pt: Optional[float] = None, bold: bool = False):
        """Copy runs from a PowerPoint paragraph, keeping bold/italic/underline/links."""
        for el in src_par._p.iterchildren():
            kind = _local(el)
            if kind == "r":
                run = PptxRun(el, src_par)
                text = clean(run.text)
                if not text:
                    continue
                font = run.font
                url = None
                try:
                    url = run.hyperlink.address
                except Exception:
                    pass
                if url:
                    self._add_hyperlink(dest_par, url, text, font.bold or bold, font.italic, size_pt)
                else:
                    r = dest_par.add_run(text)
                    r.bold = True if (font.bold or bold) else None
                    r.italic = font.italic
                    r.underline = font.underline if font.underline not in (None, False) else None
                    if size_pt:
                        r.font.size = Pt(size_pt)
            elif kind == "br":
                dest_par.add_run().add_break(WD_BREAK.LINE)
            elif kind == "fld":
                t = el.find(qn("a:t"))
                if t is not None and t.text:
                    r = dest_par.add_run(clean(t.text))
                    if bold:
                        r.bold = True
                    if size_pt:
                        r.font.size = Pt(size_pt)

    @staticmethod
    def _add_hyperlink(paragraph, url: str, text: str, bold, italic, size_pt):
        rel = paragraph.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        link = OxmlElement("w:hyperlink")
        link.set(qn("r:id"), rel)
        r_el = OxmlElement("w:r")
        link.append(r_el)
        paragraph._p.append(link)
        run = DocxRun(r_el, paragraph)
        run.text = text
        run.font.color.rgb = LINK
        run.font.underline = True
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if size_pt:
            run.font.size = Pt(size_pt)

    # ------------------------------------------------------------- pictures
    def _add_picture(self, shape):
        try:
            image = shape.image
        except Exception:
            self._placeholder_line("[Picture could not be read]")
            return

        blob, ext = image.blob, (image.ext or "").lower()
        needs_pillow = ext not in _DOCX_NATIVE_IMAGES or any(
            getattr(shape, f"crop_{side}", 0) for side in ("left", "right", "top", "bottom")
        )
        if needs_pillow:
            try:
                blob = self._prepare_with_pillow(blob, shape)
            except Exception:
                self.result.warnings.append(
                    f"Slide {self._slide_no}: picture '{shape.name}' uses an unsupported format ({ext or 'unknown'})"
                )
                self._placeholder_line(f"[Picture omitted - unsupported format: {ext or 'unknown'}]")
                return

        # Size: keep the picture's proportion of the slide width, within the page.
        share = (shape.width or self.slide_w) / self.slide_w
        width_cm = max(1.5, min(self.usable_w_cm, self.usable_w_cm * share))
        aspect = (shape.height or 1) / (shape.width or 1)
        if width_cm * aspect > self.usable_h_cm:
            width_cm = self.usable_h_cm / aspect

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = False
        p.add_run().add_picture(io.BytesIO(blob), width=Cm(width_cm))
        self.result.images += 1

    @staticmethod
    def _prepare_with_pillow(blob: bytes, shape) -> bytes:
        img = Image.open(io.BytesIO(blob))
        img.load()
        w, h = img.size
        cl, cr = getattr(shape, "crop_left", 0) or 0, getattr(shape, "crop_right", 0) or 0
        ct, cb = getattr(shape, "crop_top", 0) or 0, getattr(shape, "crop_bottom", 0) or 0
        if any((cl, cr, ct, cb)):
            box = (int(w * cl), int(h * ct), int(w * (1 - cr)), int(h * (1 - cb)))
            if box[2] > box[0] and box[3] > box[1]:
                img = img.crop(box)
        if img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
            img = img.convert("RGB")
        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()

    # --------------------------------------------------------------- tables
    def _add_table(self, ptable):
        rows, cols = len(ptable.rows), len(ptable.columns)
        if not rows or not cols:
            return
        table = self.doc.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        has_header = bool(getattr(ptable, "first_row", True))

        for r in range(rows):
            for c in range(cols):
                pcell = ptable.cell(r, c)
                if pcell.is_spanned:
                    continue
                dcell = table.cell(r, c)
                if pcell.is_merge_origin:
                    dcell = dcell.merge(table.cell(r + pcell.span_height - 1, c + pcell.span_width - 1))

                is_header = has_header and r == 0
                first = True
                for para in pcell.text_frame.paragraphs:
                    dp = dcell.paragraphs[0] if first else dcell.add_paragraph()
                    first = False
                    dp.paragraph_format.space_after = Pt(2)
                    self._write_runs(dp, para, size_pt=10, bold=is_header)
                if is_header:
                    _shade(dcell._tc.get_or_add_tcPr(), "D9E2F3")

        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        self.result.tables += 1

    # --------------------------------------------------------------- charts
    def _add_chart(self, chart):
        """Word cannot host a PowerPoint chart, so write its data as a table."""
        try:
            plot = chart.plots[0]
            categories = [clean(str(c)) for c in plot.categories]
            series = [(clean(str(s.name)), list(s.values)) for p in chart.plots for s in p.series]
            if not series:
                raise ValueError("no series")
        except Exception:
            self._placeholder_line("[Chart omitted - its data could not be read]")
            self.result.warnings.append(f"Slide {self._slide_no}: chart data could not be read")
            return

        title = ""
        try:
            if chart.has_title:
                title = clean(chart.chart_title.text_frame.text)
        except Exception:
            pass
        cap = self.doc.add_paragraph()
        r = cap.add_run(f"Chart: {title}" if title else "Chart data")
        r.bold = True
        r.font.color.rgb = ACCENT
        cap.paragraph_format.keep_with_next = True

        table = self.doc.add_table(rows=1 + len(categories), cols=1 + len(series))
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.cell(0, 0).text = ""
        for j, (name, _) in enumerate(series, start=1):
            table.cell(0, j).text = name
        for i, cat in enumerate(categories, start=1):
            table.cell(i, 0).text = cat
            for j, (_, values) in enumerate(series, start=1):
                v = values[i - 1] if i - 1 < len(values) else ""
                table.cell(i, j).text = "" if v is None else (f"{v:g}" if isinstance(v, (int, float)) else str(v))
        for c in table.rows[0].cells:
            _shade(c._tc.get_or_add_tcPr(), "D9E2F3")
            for par in c.paragraphs:
                for run in par.runs:
                    run.bold = True
        for row in table.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    par.paragraph_format.space_after = Pt(1)
                    for run in par.runs:
                        run.font.size = Pt(10)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        self.result.charts += 1

    # ---------------------------------------------------------------- notes
    def _add_notes(self, slide):
        if not slide.has_notes_slide:
            return
        text = clean(slide.notes_slide.notes_text_frame.text).strip() if slide.notes_slide.notes_text_frame else ""
        if not text:
            return
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        _shade(cell._tc.get_or_add_tcPr(), "F2F2F2")
        head = cell.paragraphs[0]
        hr = head.add_run("Speaker notes")
        hr.bold = True
        hr.font.size = Pt(9.5)
        hr.font.color.rgb = MUTED
        for line in text.splitlines():
            if not line.strip():
                continue
            p = cell.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(line)
            run.italic = True
            run.font.size = Pt(10)
            run.font.color.rgb = MUTED
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)


# --------------------------------------------------------------------------- #
# Convenience + CLI
# --------------------------------------------------------------------------- #

def unique_path(path: Path) -> Path:
    """name.docx -> name (1).docx if it already exists."""
    if not path.exists():
        return path
    n = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def convert_file(src, dst, options: Optional[ConvertOptions] = None,
                 progress: Optional[ProgressFn] = None) -> ConvertResult:
    return PptxToDocx(options).convert(src, dst, progress)


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Convert a .pptx file to .docx")
    ap.add_argument("source")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--notes", action="store_true", help="include speaker notes")
    ap.add_argument("--pages", action="store_true", help="start each slide on a new page")
    args = ap.parse_args()

    out = Path(args.output) if args.output else Path(args.source).with_suffix(".docx")
    opts = ConvertOptions(include_notes=args.notes,
                          slide_separator=SEPARATOR_PAGE if args.pages else SEPARATOR_RULE)
    res = convert_file(args.source, out, opts)
    print(f"Wrote {res.output_path}  ({res.slides_converted}/{res.slides_total} slides, "
          f"{res.images} images, {res.tables} tables, {res.charts} charts)")
    for w in res.warnings:
        print("  warning:", w)
