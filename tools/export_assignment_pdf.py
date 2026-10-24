"""Export matching PDF and Word reports with code and real notebook outputs.

Run from the repository root:
    python tools/export_assignment_pdf.py --execute

Requires the notebook's dependencies plus nbclient, nbformat, reportlab, python-docx.
Both reports use the same content and stored notebook PNGs and text.
"""

import argparse
import base64
import io
from pathlib import Path
from xml.sax.saxutils import escape

import nbformat
from export_assignment_word import write_word_report
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
NOTEBOOK = ROOT / "QC_Ekert_FA1bit_and_2bit_examp.ipynb"
PDF = ROOT / "Module4_4bit_Quantum_Adder.pdf"
WORD = ROOT / "Module4_4bit_Quantum_Adder.docx"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true", help="Run a fresh kernel first")
    args = parser.parse_args()
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
        nbformat.write(notebook, NOTEBOOK)

    for cell in notebook.cells:
        if cell.cell_type == "code":
            if not isinstance(cell.execution_count, int):
                raise ValueError("Notebook has unexecuted code; use --execute")
            if any(output.output_type == "error" for output in cell.outputs):
                raise ValueError("Notebook contains an execution error")

    def cell(index):
        tag = f"assignment-cell-{index}" if isinstance(index, int) else index
        return next(c for c in notebook.cells if tag in c.metadata.get("tags", []))

    def source(index):
        return cell(index).source.rstrip()

    def output_text(index):
        return "".join(
            output.text for output in cell(index).outputs if output.output_type == "stream"
        ).rstrip()

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
    styles.add(ParagraphStyle("CaptionQC", fontSize=9, leading=12, spaceAfter=10))
    styles.add(ParagraphStyle("CodeQC", fontName="Courier", fontSize=8, leading=9.5))
    styles.add(ParagraphStyle("OutputQC", fontName="Courier", fontSize=8, leading=10))
    styles["Title"].fontSize = 19
    styles["Title"].leading = 24
    styles["Title"].alignment = TA_CENTER
    doc = SimpleDocTemplate(
        str(PDF),
        pagesize=letter,
        leftMargin=48,
        rightMargin=48,
        topMargin=46,
        bottomMargin=42,
        title="4-bit Quantum Ripple-Carry Adder",
        subject="Module 4: Qiskit code, simulator test cases, and histograms",
        author=author["name"],
    )
    story = []

    def paragraph(text, style="BodyTextQC"):
        story.append(Paragraph(text, styles[style]))

    def heading(text):
        story.append(Paragraph(text, styles["Heading2"]))

    def code(text, output=False):
        style = styles["OutputQC" if output else "CodeQC"]
        for line in text.splitlines():
            if stringWidth(line, style.fontName, style.fontSize) > doc.width:
                raise ValueError(f"Code line would be clipped: {line}")
        story.append(Preformatted(text, style))
        story.append(Spacer(1, 8))

    def observations(tag):
        heading("Observations")
        for line in cell(tag).source.splitlines():
            if line.startswith("- "):
                story.append(Paragraph(
                    escape(line[2:]), styles["ObservationQC"], bulletText="\u2022"
                ))

    def image_from_output(index, max_width, max_height):
        png = next(
            output.data["image/png"]
            for output in cell(index).outputs
            if "image/png" in output.get("data", {})
        )
        raw = io.BytesIO(base64.b64decode(png))
        width, height = ImageReader(raw).getSize()
        scale = min(max_width / width, max_height / height)
        figure = Image(raw, width=width * scale, height=height * scale)
        figure.notebook_png = raw.getvalue()
        story.append(figure)
        story.append(Spacer(1, 10))

    def table(rows, widths):
        t = Table(rows, colWidths=widths, repeatRows=1, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8eef5")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("LEADING", (0, 0), (-1, -1), 12),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c3cbd5")),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))
        story.append(t)
        story.append(Spacer(1, 10))

    paragraph("4-bit Quantum Ripple-Carry Adder in Qiskit", "Title")
    paragraph(escape(author["name"]), "AuthorQC")
    paragraph(escape(author["affiliation"]), "AffiliationQC")
    paragraph("Module 4 assignment: extending the supplied 1-bit and 2-bit full adders")
    heading("1. Design and simulation method")
    paragraph(
        "The solution retains the class notebook's SUM, CARRY, and reverse-CARRY "
        "gate sequences. It extends the 2-bit layout to four bits using "
        "3n + 1 = 13 qubits. The operands a and b are unsigned 4-bit integers "
        "(0 to 15), with an optional carry-in of 0 or 1. The five output bits "
        "represent a + b + carry-in, including the final carry-out."
    )
    table([
        ["Bit i", "Carry c_i", "Operand a_i", "b_i becomes sum s_i"],
        ["0", "q0 (carry-in)", "q1", "q2"],
        ["1", "q3", "q4", "q5"],
        ["2", "q6", "q7", "q8"],
        ["3", "q9", "q10", "q11"],
        ["Final carry", "q12 = s4", "", ""],
    ], [75, 127, 119, 195])
    paragraph(
        "First, four CARRY blocks propagate the carry upward. At the highest bit, "
        "a CNOT restores b3 before SUM writes s3. Then reverse-CARRY and SUM "
        "run downward for bits 2, 1, and 0. This restores the internal carry "
        "qubits q3, q6, and q9 to zero while preserving a and the initial carry-in. "
        "The arithmetic circuit contains 14 Toffoli (CCX) and 16 CNOT (CX) gates."
    )
    paragraph(
        "Sum bit si is measured into classical bit i, and q12 into bit 4. "
        "Qiskit displays bit 4 on the left, so labels read s4s3s2s1s0. "
        "For example, 1011 + 0110 = 10001 (11 + 6 = 17)."
    )
    paragraph(
        "All measured results use the local ideal Qiskit Aer SamplerV2, "
        "1,024 shots per circuit, seed 4, and no noise model. Counts come "
        "from get_counts(); expected sums are computed separately for comparison. "
        "These are simulator results. The exhaustive check uses unmeasured "
        "Qiskit Statevector circuits."
    )
    heading("Executed notebook environment")
    code(source(1))
    code(output_text(1), output=True)
    story.append(PageBreak())

    heading("2. Code: imports and original arithmetic blocks")
    code(source(2))
    code(source(3))
    heading("General adder: input preparation")
    split_marker = "    # QC carry ... (ripple up)"
    preparation, remainder = source(11).split(split_marker, 1)
    code(preparation.rstrip())
    story.append(PageBreak())

    heading("General adder continued: carry, sum, and measurement")
    code(split_marker + remainder)
    heading("3. Example: binary 1011 + 0110 = 10001")
    code(source(13))
    code(output_text(13), output=True)
    code(source(14))
    code(output_text(14), output=True)
    paragraph(
        "The notebook also displays qc.draw(). The following page shows the "
        "histogram saved by the executed example cell.", "CaptionQC"
    )
    story.append(PageBreak())

    heading("Example histogram: 11 + 6 = 17")
    image_from_output(14, doc.width, 260)
    observations("histogram-observations-example")
    story.append(PageBreak())

    heading("4. Selected test cases")
    table([
        ["a (4 bits)", "b (4 bits)", "Carry-in", "Purpose"],
        ["0000 (0)", "0000 (0)", "0", "Zero inputs"],
        ["0101 (5)", "0010 (2)", "0", "No carry generated"],
        ["0111 (7)", "0001 (1)", "0", "Carry propagation; no overflow"],
        ["1001 (9)", "0111 (7)", "0", "Overflow into fifth output bit"],
        ["1111 (15)", "1111 (15)", "0", "Maximum operands"],
        ["1111 (15)", "0000 (0)", "1", "Carry-in propagates through all bits"],
    ], [85, 85, 65, 281])
    paragraph(
        "On this ideal simulator, basis-state inputs give deterministic outputs. "
        "The assertions require the entire counts dictionary to equal the "
        "expected five-bit result with 1,024 counts."
    )
    heading("Test-case and histogram code")
    code(source(16))
    heading("Actual simulator test output")
    code(output_text(16), output=True)
    story.append(PageBreak())

    for index in range(6):
        tag = f"histogram-test-{index}"
        heading(escape(cell(tag).metadata["report_title"]))
        code(source(tag))
        image_from_output(tag, doc.width, 300)
        observations(f"histogram-observations-test-{index}")
        story.append(PageBreak())

    heading("5. Exhaustive verification")
    paragraph(
        "The following code checks every 16 x 16 x 2 = 512 combination of a, b, "
        "and carry-in. It verifies the five-bit sum, preservation of a and "
        "carry-in, and reset of the three internal carry qubits. It also checks "
        "that the output is one basis state with probability 1."
    )
    code(source(18))
    code(output_text(18), output=True)
    paragraph(
        "The example, all six selected measured cases, and all 512 exhaustive "
        "input combinations pass. This verifies the 4-bit adder for every "
        "allowed unsigned input and both carry-in values under ideal simulation."
    )
    heading("Sources")
    paragraph(
        "Supplied class notebook: QC_Ekert_FA1bit_and_2bit_examp.ipynb. "
        "The SUM, CARRY, and reverse-CARRY definitions are retained unchanged.<br/>"
        'Vedral, Barenco, and Ekert, <i>Quantum Networks for Elementary Arithmetic '
        'Operations</i>, 1995 preprint (journal publication: Phys. Rev. A 54, 147, 1996): '
        '<link href="https://arxiv.org/abs/quant-ph/9511018" color="blue">arXiv paper</link>.<br/>'
        '<link href="https://quantum.cloud.ibm.com/docs/en/guides/bit-ordering" '
        'color="blue">IBM Qiskit bit-ordering documentation</link>; '
        '<link href="https://qiskit.github.io/qiskit-aer/stubs/qiskit_aer.primitives.SamplerV2.html" '
        'color="blue">Qiskit Aer SamplerV2 reference</link>.',
        "CaptionQC",
    )

    write_word_report(story, WORD, author["name"])
    doc.build(story)
    print(f"Wrote {PDF}")
    print(f"Wrote {WORD}")


if __name__ == "__main__":
    main()
