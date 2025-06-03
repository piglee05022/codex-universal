from typing import Dict

try:
    from docx import Document
    from docx.shared import Pt
except ImportError:  # pragma: no cover - docx may not be installed
    Document = None
    Pt = None

class LegalDocumentGenerator:
    """Generate legal filings in Traditional Chinese."""

    def __init__(self, case_info: Dict[str, any]):
        self.case_info = case_info
        if Document is None:
            raise RuntimeError(
                "python-docx is required to generate documents. Please install it via 'pip install python-docx'."
            )
        self.doc = Document()
        style = self.doc.styles['Normal']
        font = style.font
        font.name = '標楷體'
        font.size = Pt(16)

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


def create_filing(case_info: Dict[str, any], output_path: str) -> None:
    """Helper function to quickly generate a filing document."""
    generator = LegalDocumentGenerator(case_info)
    generator.build()
    generator.save(output_path)

