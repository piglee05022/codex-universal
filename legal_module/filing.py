"""Generate Traditional Chinese legal documents using python-docx."""

from typing import Dict

try:  # pragma: no cover - optional dependency may be missing
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_LINE_SPACING
except ImportError:  # pragma: no cover - docx may not be installed
    Document = None
    Pt = None
    Cm = None
    WD_LINE_SPACING = None

try:  # pragma: no cover - optional dependency
    from docx2pdf import convert as docx2pdf_convert
except Exception:  # pragma: no cover - docx2pdf is truly optional
    docx2pdf_convert = None

class LegalDocumentGenerator:
    """Generate legal filings in Traditional Chinese."""

    def __init__(self, case_info: Dict[str, any]):
        self.case_info = case_info
        if Document is None:
            raise RuntimeError(
                "python-docx is required to generate documents. Please install it via 'pip install python-docx'."
            )
        self.doc = Document()
        style = self.doc.styles["Normal"]
        font = style.font
        font.name = "標楷體"
        font.size = Pt(16)
        style.paragraph_format.line_spacing = 1.5
        for section in self.doc.sections:
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)

    def build(self) -> None:
        self.doc.add_heading(self.case_info.get('title', '起訴狀'), level=1)
        self._add_basic_info()
        self._add_claims()
        self._add_facts()
        self._add_laws()
        self._add_evidence()

    def save(self, filepath: str) -> None:
        self.doc.save(filepath)

    # Internal helpers
    def _add_basic_info(self) -> None:
        self.doc.add_paragraph(f"案號：{self.case_info.get('case_number', '')}")
        self.doc.add_paragraph(f"當事人：{self.case_info.get('parties', '')}")
        self.doc.add_paragraph(f"法院：{self.case_info.get('court', '')}")

    def _add_claims(self) -> None:
        self.doc.add_paragraph("壹、訴之聲明")
        self.doc.add_paragraph(self.case_info.get('claims', ''))

    def _add_facts(self) -> None:
        self.doc.add_paragraph("貳、事實與理由")
        self.doc.add_paragraph(self.case_info.get('facts', ''))

    def _add_laws(self) -> None:
        self.doc.add_paragraph("參、法律依據")
        for law in self.case_info.get('laws', []):
            self.doc.add_paragraph(f"• {law}")

    def _add_evidence(self) -> None:
        self.doc.add_paragraph("肆、證據目錄")
        for ev in self.case_info.get('evidence', []):
            self.doc.add_paragraph(f"【{ev['id']}】{ev['summary']}")


def create_legal_filing(case_info: Dict[str, any], output_path: str, *, export_pdf: bool = False) -> None:
    """Create a legal filing as a Word document.

    Parameters
    ----------
    case_info:
        Case details used to populate the document.
    output_path:
        Path to the generated Word document.
    export_pdf:
        If ``True`` and ``docx2pdf`` is installed, also export a PDF.
    """

    generator = LegalDocumentGenerator(case_info)
    generator.build()
    generator.save(output_path)

    if export_pdf and docx2pdf_convert is not None:
        docx2pdf_convert(output_path)

