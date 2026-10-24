"""Export the adder code, circuit diagrams, and five tests to PDF and Word.

Run from the repository root:
    python tools/export_assignment_pdf.py --execute
"""

import argparse
import base64
import io
from pathlib import Path
from xml.sax.saxutils import escape

import nbformat
from export_assignment_word import write_word_report
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    Image, PageBreak, Paragraph, Preformatted, SimpleDocTemplate, Spacer,
)


ROOT = Path(__file__).resolve().parents[1]
NOTEBOOK = ROOT / "QC_Ekert_FA1bit_and_2bit_examp.ipynb"
PDF = ROOT / "Module4_4bit_Quantum_Adder.pdf"
WORD = ROOT / "Module4_4bit_Quantum_Adder.docx"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true", help="Execute the notebook first")
    parser.add_argument("--output-dir", type=Path, default=ROOT)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = args.output_dir / PDF.name
    word_path = args.output_dir / WORD.name
    notebook = nbformat.read(NOTEBOOK, as_version=4)
    author = notebook.metadata["authors"][0]
    if args.execute:
        import asyncio
        import sys

        from nbclient import NotebookClient

        if sys.platform == "win32":
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        NotebookClient(
            notebook,
            timeout=180,
            kernel_name="python3",
            resources={"metadata": {"path": str(ROOT)}},
        ).execute()
        with NOTEBOOK.open("w", encoding="utf-8", newline="\n") as output:
            nbformat.write(notebook, output)

    for cell in notebook.cells:
        if cell.cell_type == "code":
            if not isinstance(cell.execution_count, int):
                raise ValueError("Notebook has unexecuted code; use --execute")
            if any(output.output_type == "error" for output in cell.outputs):
                raise ValueError("Notebook contains an execution error")

    def cell(tag):
        if isinstance(tag, int):
            tag = f"assignment-cell-{tag}"
        return next(c for c in notebook.cells if tag in c.metadata.get("tags", []))

    def source(tag):
        return cell(tag).source.rstrip()

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("BodyTextQC", fontSize=10, leading=14, spaceAfter=9))
    styles.add(ParagraphStyle(
        "AuthorQC", fontSize=11, leading=15, alignment=TA_CENTER, spaceAfter=2
    ))
    styles.add(ParagraphStyle(
        "AffiliationQC", fontSize=10, leading=13, alignment=TA_CENTER, spaceAfter=12
    ))
    styles.add(ParagraphStyle(
        "ObservationQC", parent=styles["BodyTextQC"], leftIndent=12, bulletIndent=0
    ))
    styles.add(ParagraphStyle("CodeQC", fontName="Courier", fontSize=8, leading=9.5))
    styles.add(ParagraphStyle("OutputQC", fontName="Courier", fontSize=8, leading=10))
    styles["Title"].fontSize = 19
    styles["Title"].leading = 24
    styles["Title"].alignment = TA_CENTER
    document = SimpleDocTemplate(
        str(pdf_path), pagesize=letter,
        leftMargin=48, rightMargin=48, topMargin=46, bottomMargin=42,
        title="4-Bit Adder in Qiskit",
        subject="Module 4: four-bit adder code, test cases, and histograms",
        author=author["name"],
    )
    story = []

    def paragraph(text, style="BodyTextQC"):
        story.append(Paragraph(text, styles[style]))

    def heading(text):
        paragraph(text, "Heading2")

    def code(text, output=False):
        style = styles["OutputQC" if output else "CodeQC"]
        for line in text.splitlines():
            if stringWidth(line, style.fontName, style.fontSize) > document.width:
                raise ValueError(f"Code line would be clipped: {line}")
        story.append(Preformatted(text, style))
        story.append(Spacer(1, 8))

    def figure(tag, max_height):
        png = next(
            item.data["image/png"] for item in cell(tag).outputs
            if "image/png" in item.get("data", {})
        )
        raw = io.BytesIO(base64.b64decode(png))
        width, height = ImageReader(raw).getSize()
        scale = min(document.width / width, max_height / height)
        diagram = Image(raw, width=width * scale, height=height * scale)
        diagram.notebook_png = raw.getvalue()
        story.append(diagram)
        story.append(Spacer(1, 10))

    paragraph("4-Bit Adder in Qiskit", "Title")
    paragraph(escape(author["name"]), "AuthorQC")
    paragraph(escape(author["affiliation"]), "AffiliationQC")
    paragraph(
        "This circuit extends the SUM, CARRY, and reverse-CARRY design in "
        "QC_Ekert_FA1bit_and_2bit_examp.ipynb from two bits to four bits. "
        "The initial carry is zero. Five measured bits represent the sum, "
        "including the final carry-out."
    )
    heading("4-Bit Adder Code")
    code(source(2))
    code(source(3))
    split_marker = "    # Compute the four carries."
    preparation, remainder = source(11).split(split_marker, 1)
    code(preparation.rstrip())
    story.append(PageBreak())
    heading("4-Bit Adder Code (continued)")
    code(split_marker + remainder)
    heading("Test Cases")
    paragraph(
        "The five tests use four-bit binary operands and 1,024 shots per "
        "circuit on Qiskit Aer. Each assertion compares the measured counts "
        "with the expected five-bit sum."
    )
    code(source(16))
    heading("Test Results")
    output = "".join(
        item.text for item in cell(16).outputs if item.output_type == "stream"
    ).rstrip()
    code(output, output=True)

    for index in range(5):
        story.append(PageBreak())
        tag = f"histogram-test-{index}"
        heading(escape(cell(tag).metadata["report_title"]) + " (Circuit)")
        circuit_tag = f"circuit-test-{index}"
        code(source(circuit_tag))
        figure(circuit_tag, 575)
        story.append(PageBreak())
        heading(escape(cell(tag).metadata["report_title"]))
        code(source(tag))
        figure(tag, 300)
        heading("Observations")
        for line in cell(f"histogram-observations-test-{index}").source.splitlines():
            if line.startswith("- "):
                story.append(Paragraph(
                    escape(line[2:]), styles["ObservationQC"], bulletText="\u2022"
                ))

    try:
        write_word_report(story, word_path, author["name"])
    except PermissionError:
        raise SystemExit(
            f"Close {word_path.name} in Word, then run the export again. "
            "No additional report copy was created."
        ) from None
    document.build(story)
    print(f"Wrote {pdf_path}")
    print(f"Wrote {word_path}")


if __name__ == "__main__":
    main()
