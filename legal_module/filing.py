"""Generate Traditional Chinese legal documents using python-docx."""

from typing import List, Tuple, Union, TypedDict
from enum import Enum


class EvidenceEntry(TypedDict):
    """A single evidence item."""

    id: str
    summary: str


class CaseInfo(TypedDict, total=False):
    """Information required for generating a legal document."""

    title: str
    case_number: str
    parties: str
    court: str
    claims: str
    facts: str
    laws: List[str]
    evidence: List[EvidenceEntry]

    filing_type: str  # optional; defaults to ``FilingType.COMPLAINT``

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

CHINESE_NUMERALS = {
    1: "壹",
    2: "貳",
    3: "參",
    4: "肆",
    5: "伍",
    6: "陸",
    7: "柒",
    8: "捌",
    9: "玖",
    10: "拾",
}


class FilingType(str, Enum):
    """Common filing types."""

    COMPLAINT = "起訴狀"
    DEFENSE = "答辯狀"
    APPEAL = "上訴狀"
    ENFORCEMENT = "強制執行聲請狀"
    SUPPLEMENT = "補充說明狀"
    INVESTIGATION = "調查證據聲請狀"

class LegalDocumentGenerator:
    """Generate legal filings in Traditional Chinese."""

    def __init__(self, case_info: CaseInfo):
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
        title = self.case_info.get("filing_type") or self.case_info.get("title")
        if not title:
            title = FilingType.COMPLAINT.value
        self.doc.add_heading(str(title), level=1)
        self._add_basic_info()
        for idx, section in enumerate(
            [
                ("訴之聲明", self.case_info.get("claims", "")),
                ("事實與理由", self.case_info.get("facts", "")),
                ("法律依據", self.case_info.get("laws", [])),
                ("證據目錄", self.case_info.get("evidence", [])),
            ],
            start=1,
        ):
            label = CHINESE_NUMERALS.get(idx, str(idx)) + "、" + section[0]
            self.doc.add_paragraph(label)
            if section[0] == "法律依據":
                for law in section[1]:
                    self.doc.add_paragraph(f"• {law}")
            elif section[0] == "證據目錄":
                for ev in section[1]:
                    self.doc.add_paragraph(f"【{ev['id']}】{ev['summary']}")
            else:
                self.doc.add_paragraph(section[1])

    def save(self, filepath: str) -> None:
        self.doc.save(filepath)

    # Internal helpers
    def _add_basic_info(self) -> None:
        self.doc.add_paragraph(f"案號：{self.case_info.get('case_number', '')}")
        self.doc.add_paragraph(f"當事人：{self.case_info.get('parties', '')}")
        self.doc.add_paragraph(f"法院：{self.case_info.get('court', '')}")



def create_legal_filing(
    case_info: CaseInfo,
    output_path: str = "法律文書_起訴狀.docx",
    *,
    export_pdf: bool = False,
) -> Union[str, Tuple[str, str]]:
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

    if export_pdf:
        if docx2pdf_convert is None:
            raise RuntimeError(
                "docx2pdf is required for exporting PDF. Install it via 'pip install docx2pdf'."
            )
        pdf_path = output_path.rsplit(".", 1)[0] + ".pdf"
        docx2pdf_convert(output_path, pdf_path)
        return output_path, pdf_path

    return output_path

