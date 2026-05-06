#!/usr/bin/env python3
"""
mcq_transfer.py — MCQ Bank Transfer Tool
=========================================
Copies every multiple-choice question from a PEA-format source document into
an Arabic Question-Bank template document.

Handles:
  • Plain text (any font / formatting)
  • OMML math equations  (copied as raw XML – equations never break)
  • Embedded images / charts (relationship IDs are re-mapped automatically)
  • All four options  (A→أ  B→ب  C→ت  D→ث)
  • Correct-answer star (*) placed in the star column
  • Unlimited questions  (template block is cloned for each question)

Usage:
    python mcq_transfer.py          # opens the GUI
"""

import os
import re
import threading
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

# English option letter → Arabic option letter used in the template
OPTION_MAP: dict = {
    'A': 'أ',
    'B': 'ب',
    'C': 'ت',
    'D': 'ث',
}
OPTION_MAP_REV: dict = {v: k for k, v in OPTION_MAP.items()}

# XML namespace URIs
NS_MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS_WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_REL  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMBED_ATTR = f'{{{NS_REL}}}embed'

# Lines in the source document to silently ignore
SKIP_PHRASES = (
    "REFER TO THE FOLLOWING",
    "أسئلة الاختيار",
    "فيما يلي سؤال",
    "أسئلة المقارنة",
    "EXHIBIT",
)


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 1 – PARSE SOURCE DOCUMENT
# ─────────────────────────────────────────────────────────────────────────────

def extract_exam_key(doc) -> dict:
    """
    Read the EXAMKEY table (first table in the source document).

    Expected column layout (0-indexed):
        col 0 – sequential number  (1, 2, 3 …)
        col 1 – question ID        (e.g. "TSQ10084-01")
        col 4 – correct answer     (A / B / C / D)

    Returns
    -------
    dict  {question_number (int) →
               {'id': str, 'answer': str (Arabic letter, e.g. 'ب')}}
    """
    if not doc.tables:
        raise ValueError(
            "Source document has no tables.\n"
            "Expected the EXAMKEY as the very first table."
        )

    key_data: dict = {}
    for row in doc.tables[0].rows[1:]:      # row[0] = header – skip it
        cells = row.cells
        if len(cells) < 5:
            continue
        seq      = cells[0].text.strip()
        q_id     = cells[1].text.strip()
        ans_eng  = cells[4].text.strip().upper()
        if seq.isdigit():
            key_data[int(seq)] = {
                'id':     q_id,
                'answer': OPTION_MAP.get(ans_eng, ''),
            }
    return key_data


def _para_has_content(para) -> bool:
    """Return True when a paragraph carries visible or embedded content."""
    if para.text.strip():
        return True
    el = para._element
    has_math    = bool(el.xpath('.//m:oMath',   namespaces={'m': NS_MATH}))
    has_drawing = bool(el.xpath('.//w:drawing', namespaces={'w': NS_WORD}))
    return has_math or has_drawing


def parse_questions(doc) -> dict:
    """
    Walk every paragraph in the source document and group them into questions.

    A question block begins at "QUESTION N" and contains:
        • body paragraphs  (everything before the first option)
        • option A … D paragraphs  (everything under each "X. …" marker)

    Returns
    -------
    dict  {question_number (int) →
               {'body': [...], 'A': [...], 'B': [...], 'C': [...], 'D': [...]}}
    """
    questions:       dict = {}
    current_q:       int  = None   # active question number
    current_section: str  = None   # 'body' | 'A' | 'B' | 'C' | 'D'

    re_question = re.compile(r'^QUESTION\s+(\d+)', re.IGNORECASE)
    re_option   = re.compile(r'^([A-D])\.\s*(.*)', re.DOTALL | re.IGNORECASE)

    for para in doc.paragraphs:
        text = para.text.strip()

        # Completely blank paragraph with no embedded content → skip
        if not text and not _para_has_content(para):
            continue

        # ── New QUESTION header ──────────────────────────────────────────────
        m = re_question.match(text)
        if m:
            current_q      = int(m.group(1))
            current_section = 'body'
            questions[current_q] = {
                'body': [], 'A': [], 'B': [], 'C': [], 'D': [],
            }
            continue

        # Nothing active yet
        if current_q is None:
            continue

        # Boilerplate / instructional lines – discard
        if any(phrase in text for phrase in SKIP_PHRASES):
            continue

        # ── Option header  (A.  /  B.  /  C.  /  D.) ────────────────────────
        m = re_option.match(text)
        if m:
            letter = m.group(1).upper()
            current_section = letter
            questions[current_q][letter].append(para)
            continue

        # ── Regular content ──────────────────────────────────────────────────
        if current_section:
            questions[current_q][current_section].append(para)

    return questions


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 2 – XML SURGERY  (text / math / images)
# ─────────────────────────────────────────────────────────────────────────────

def _transfer_image_rels(p_xml, source_doc, target_doc):
    """
    For every r:embed attribute inside p_xml:
      1. Locate the image Part in source_doc.
      2. Add it to target_doc (returns a fresh rId).
      3. Rewrite the attribute in-place so the image resolves correctly.
    """
    for elem in p_xml.iter():
        if EMBED_ATTR in elem.attrib:
            old_rid = elem.attrib[EMBED_ATTR]
            try:
                rel     = source_doc.part.rels[old_rid]
                new_rid = target_doc.part.relate_to(rel.target_part, rel.reltype)
                elem.attrib[EMBED_ATTR] = new_rid
            except (KeyError, AttributeError):
                pass    # best-effort; image may still render via the old rId


def inject_paragraphs_into_cell(source_paras, target_cell, source_doc, target_doc):
    """
    Replace the contents of target_cell with the paragraphs in source_paras.

    Each source paragraph is deep-copied at the XML level, which preserves:
      • character formatting (bold, italic, font, colour …)
      • OMML math equations  (the <m:oMath> subtree is copied verbatim)
      • Inline / floating images  (relationships are re-mapped to target_doc)
    """
    tc = target_cell._tc

    # Clear existing cell content
    for node in list(tc.findall(qn('w:p'))):
        tc.remove(node)
    for node in list(tc.findall(qn('w:tbl'))):
        tc.remove(node)

    real_paras = [p for p in source_paras if _para_has_content(p)]

    if not real_paras:
        tc.append(OxmlElement('w:p'))   # Word requires ≥1 paragraph per cell
        return

    for src in real_paras:
        new_p = deepcopy(src._element)
        _transfer_image_rels(new_p, source_doc, target_doc)
        tc.append(new_p)


def _set_cell_text(cell, text: str):
    """Overwrite a cell with a single plain-text string."""
    tc = cell._tc
    for node in list(tc.findall(qn('w:p'))):
        tc.remove(node)
    new_p = OxmlElement('w:p')
    tc.append(new_p)
    if text:
        run = OxmlElement('w:r')
        t   = OxmlElement('w:t')
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        run.append(t)
        new_p.append(run)


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 3 – TEMPLATE STRUCTURE DETECTION
# ─────────────────────────────────────────────────────────────────────────────

class TemplateStructure:
    """
    Holds the cell positions (relative to the start of one question block)
    for every field that needs to be filled.
    """
    def __init__(self):
        self.id_cell:    tuple = None   # (rel_row, col_index)
        self.body_cell:  tuple = None
        # Arabic letter → {'content': (rel_row, col), 'star': (rel_row, col)}
        self.options:    dict  = {}
        self.block_rows: int   = 0      # rows that make up one question block

    def is_valid(self) -> bool:
        return (
            self.id_cell   is not None and
            self.body_cell is not None and
            len(self.options) >= 2
        )

    def summary(self) -> str:
        lines = [
            f"  Block rows  : {self.block_rows}",
            f"  ID cell     : row {self.id_cell[0]}, col {self.id_cell[1]}",
            f"  Body cell   : row {self.body_cell[0]}, col {self.body_cell[1]}",
        ]
        for letter, info in self.options.items():
            cr, cc = info['content']
            sr, sc = info['star']
            lines.append(
                f"  Option '{letter}' : content(row {cr}, col {cc})  "
                f"star(row {sr}, col {sc})"
            )
        return "\n".join(lines)


def _cell_text_xml(tc_elem) -> str:
    """Extract all text from a <w:tc> XML element."""
    return ''.join(
        (t.text or '') for t in tc_elem.findall('.//' + qn('w:t'))
    ).strip()


def detect_template_structure(table) -> TemplateStructure:
    """
    Auto-detect cell positions by scanning the first N rows for known
    Arabic keyword labels and option letters.

    Labels recognised:
        ID cell    ← rows containing  'رمز السؤال'
        Body cell  ← rows containing  'نص السؤال'
        Options    ← rows containing  أ / ب / ت / ث
    """
    tbl_elem  = table._tbl
    rows_xml  = tbl_elem.findall(qn('w:tr'))

    ARABIC_OPTIONS = {'أ', 'ب', 'ت', 'ث'}
    ID_KEYWORDS    = {'رمز السؤال'}
    BODY_KEYWORDS  = {'نص السؤال'}

    ts = TemplateStructure()
    last_option_row = 0

    for r_idx, row_xml in enumerate(rows_xml):
        cells_xml = row_xml.findall(qn('w:tc'))

        for c_idx, cell_xml in enumerate(cells_xml):
            text = _cell_text_xml(cell_xml)

            # ── Question-ID label ────────────────────────────────────────────
            if any(kw in text for kw in ID_KEYWORDS) and ts.id_cell is None:
                # The value goes in the next column that doesn't repeat the label
                for nc in range(c_idx + 1, len(cells_xml)):
                    nc_text = _cell_text_xml(cells_xml[nc])
                    if not any(kw in nc_text for kw in ID_KEYWORDS):
                        ts.id_cell = (r_idx, nc)
                        break
                # Fallback: same cell (unusual, but handled gracefully)
                if ts.id_cell is None:
                    ts.id_cell = (r_idx, c_idx)

            # ── Question-body label ──────────────────────────────────────────
            elif any(kw in text for kw in BODY_KEYWORDS) and ts.body_cell is None:
                if c_idx + 1 < len(cells_xml):
                    ts.body_cell = (r_idx, c_idx + 1)
                else:
                    # Label spans full width → body content goes into this cell
                    ts.body_cell = (r_idx, c_idx)

            # ── Arabic option letter ─────────────────────────────────────────
            elif text in ARABIC_OPTIONS and text not in ts.options:
                content_col = c_idx + 1 if c_idx + 1 < len(cells_xml) else c_idx
                # Star column = last physical column in the row
                star_col = len(cells_xml) - 1
                # Avoid content_col == star_col for very narrow tables
                if content_col == star_col and content_col > c_idx:
                    star_col = content_col
                ts.options[text] = {
                    'content': (r_idx, content_col),
                    'star':    (r_idx, star_col),
                }
                last_option_row = r_idx

    ts.block_rows = last_option_row + 1
    return ts


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 4 – CELL ACCESS  (XML-level, merged-cell-safe)
# ─────────────────────────────────────────────────────────────────────────────

def _get_cell(table, abs_row: int, col: int):
    """
    Return a docx _Cell at the given absolute row and physical column index,
    bypassing python-docx's merged-cell remapping.

    Returns None if the index is out of range.
    """
    from docx.table import _Cell
    rows = table._tbl.findall(qn('w:tr'))
    if abs_row >= len(rows):
        return None
    row_cells = rows[abs_row].findall(qn('w:tc'))
    if col >= len(row_cells):
        return None
    return _Cell(row_cells[col], table)


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 5 – FILL ONE QUESTION BLOCK
# ─────────────────────────────────────────────────────────────────────────────

def fill_question_block(
    table,
    row_offset: int,
    ts: TemplateStructure,
    q_num: int,
    k_data: dict,
    q_data: dict,
    source_doc,
    target_doc,
):
    """
    Inject question data into the block that begins at row_offset.

    Parameters
    ----------
    table       : the target Document's table
    row_offset  : first row of this question's block (0-based)
    ts          : detected template structure
    q_num       : question sequence number
    k_data      : {'id': str, 'answer': Arabic letter}
    q_data      : {'body': [...], 'A': [...], 'B': [...], 'C': [...], 'D': [...]}
    source_doc  : source Document (needed for image-rel transfer)
    target_doc  : target Document (needed for image-rel transfer)
    """
    def cell(rel_row, col):
        return _get_cell(table, row_offset + rel_row, col)

    # 1. Question ID
    if ts.id_cell:
        c = cell(*ts.id_cell)
        if c:
            _set_cell_text(c, k_data.get('id', ''))

    # 2. Question body
    if ts.body_cell:
        c = cell(*ts.body_cell)
        if c:
            inject_paragraphs_into_cell(
                q_data.get('body', []), c, source_doc, target_doc
            )

    # 3. Options + correct-answer star
    correct_arabic = k_data.get('answer', '')

    for ara_letter, info in ts.options.items():
        eng_letter    = OPTION_MAP_REV.get(ara_letter, '')
        option_paras  = q_data.get(eng_letter, [])

        # Option content
        c_content = cell(*info['content'])
        if c_content:
            if option_paras:
                inject_paragraphs_into_cell(
                    option_paras, c_content, source_doc, target_doc
                )
            else:
                _set_cell_text(c_content, '')   # clear template placeholder

        # Star marker (correct answer)
        c_star = cell(*info['star'])
        if c_star:
            _set_cell_text(c_star, '*' if ara_letter == correct_arabic else '')


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 6 – TEMPLATE BLOCK CLONING
# ─────────────────────────────────────────────────────────────────────────────

def _clone_block(table, num_rows: int) -> list:
    """
    Deep-copy the first `num_rows` rows of `table`.
    Returns a list of detached <w:tr> XML elements.
    """
    all_rows = table._tbl.findall(qn('w:tr'))
    return [
        deepcopy(all_rows[i])
        for i in range(min(num_rows, len(all_rows)))
    ]


def _append_block(table, row_elements: list):
    """Append detached <w:tr> elements to the end of `table`."""
    tbl = table._tbl
    for tr in row_elements:
        tbl.append(tr)


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 7 – MASTER ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

def process_and_transfer(
    source_path:   str,
    template_path: str,
    output_path:   str,
    log=None,
):
    """
    Full pipeline:
        load source → extract key → parse questions →
        detect template → fill blocks → save output

    Parameters
    ----------
    source_path   : .docx  PEA source file
    template_path : .docx  Arabic question-bank template
    output_path   : .docx  destination for the filled bank
    log           : callable(str) for status messages, or None
    """
    def _log(msg: str):
        if log:
            log(msg)

    # ── Load ──────────────────────────────────────────────────────────────────
    _log("Loading source document …")
    source_doc  = Document(source_path)

    _log("Loading template document …")
    target_doc  = Document(template_path)

    # ── Extract exam key ──────────────────────────────────────────────────────
    _log("Extracting exam key (question IDs & correct answers) …")
    exam_key = extract_exam_key(source_doc)
    if not exam_key:
        raise ValueError(
            "No data found in the EXAMKEY table.\n"
            "Check that the first table in the source document contains "
            "sequential numbers in column 0, question IDs in column 1, "
            "and the correct answer (A/B/C/D) in column 4."
        )
    _log(f"  → {len(exam_key)} entries in exam key.")

    # ── Parse questions ───────────────────────────────────────────────────────
    _log("Parsing questions from source document …")
    parsed_qs = parse_questions(source_doc)
    if not parsed_qs:
        raise ValueError(
            "No questions were found in the source document.\n"
            "Questions must begin with 'QUESTION 1', 'QUESTION 2', etc."
        )
    _log(f"  → {len(parsed_qs)} questions parsed.")

    # ── Detect template structure ─────────────────────────────────────────────
    if not target_doc.tables:
        raise ValueError("Template document contains no tables.")

    target_table = target_doc.tables[0]
    _log("Detecting template structure …")
    ts = detect_template_structure(target_table)

    if not ts.is_valid():
        raise ValueError(
            "Could not auto-detect the template layout.\n\n"
            "Make sure the template's first table contains at least:\n"
            "  • a cell with 'رمز السؤال'  (question ID label)\n"
            "  • a cell with 'نص السؤال'   (body label)\n"
            "  • cells with  أ  ب  ت  ث    (option labels)\n"
        )
    _log("  Template structure detected:")
    for line in ts.summary().splitlines():
        _log(line)

    # ── Determine which questions to transfer ─────────────────────────────────
    all_nums = sorted(set(exam_key) & set(parsed_qs))
    if not all_nums:
        raise ValueError(
            "The question numbers in the exam key do not match the "
            "question numbers parsed from the body of the document.\n"
            f"Exam key has: {sorted(exam_key.keys())}\n"
            f"Parsed body has: {sorted(parsed_qs.keys())}"
        )
    total = len(all_nums)
    _log(f"Transferring {total} question(s) …")

    # ── Fill blocks ───────────────────────────────────────────────────────────
    for idx, q_num in enumerate(all_nums):

        # For every question after the first, clone the template block
        if idx > 0:
            cloned = _clone_block(target_table, ts.block_rows)
            _append_block(target_table, cloned)

        row_offset = idx * ts.block_rows
        fill_question_block(
            target_table,
            row_offset,
            ts,
            q_num,
            exam_key[q_num],
            parsed_qs[q_num],
            source_doc,
            target_doc,
        )
        _log(
            f"  [{idx + 1}/{total}]  Q{q_num} → "
            f"{exam_key[q_num]['id']}  "
            f"(correct: {exam_key[q_num]['answer']})"
        )

    # ── Save ──────────────────────────────────────────────────────────────────
    _log(f"Saving output → {output_path}")
    target_doc.save(output_path)
    _log("Done!  ✓")


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 8 – GUI
# ─────────────────────────────────────────────────────────────────────────────

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    _TKINTER_AVAILABLE = True
except ImportError:
    _TKINTER_AVAILABLE = False


class TransferApp:
    """Main application window."""

    _BG      = "#ecf0f1"
    _HEADER  = "#2c3e50"
    _GREEN   = "#27ae60"
    _GREY    = "#7f8c8d"
    _LOG_BG  = "#1a1a2e"
    _LOG_FG  = "#00ff88"

    def __init__(self, root):
        self.root = root
        self.root.title("MCQ Bank Transfer Tool")
        self.root.geometry("640x520")
        self.root.resizable(False, False)
        self.root.configure(bg=self._BG)

        self.source_var   = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var   = tk.StringVar()

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=self._HEADER)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr, text="MCQ Bank Transfer Tool",
            font=("Segoe UI", 17, "bold"), fg="white", bg=self._HEADER, pady=10,
        ).pack()
        tk.Label(
            hdr, text="PEA Format  →  Arabic Question Bank",
            font=("Segoe UI", 9), fg="#bdc3c7", bg=self._HEADER, pady=(0, 8),
        ).pack()

        # File selection
        fbox = tk.LabelFrame(
            self.root, text="  Files  ",
            font=("Segoe UI", 9, "bold"), bg=self._BG, padx=14, pady=12,
        )
        fbox.pack(fill=tk.X, padx=18, pady=12)

        self._file_row(fbox, "1.  Source  (PEA format):",
                       self.source_var,   self._browse_source,   row=0)
        self._file_row(fbox, "2.  Question Bank Template:",
                       self.template_var, self._browse_template, row=1)
        self._file_row(fbox, "3.  Save output as:",
                       self.output_var,   self._browse_output,   row=2)

        # Progress / log
        pbox = tk.LabelFrame(
            self.root, text="  Log  ",
            font=("Segoe UI", 9, "bold"), bg=self._BG, padx=14, pady=10,
        )
        pbox.pack(fill=tk.X, padx=18)

        self.progress_bar = ttk.Progressbar(
            pbox, mode='indeterminate', length=580,
        )
        self.progress_bar.pack(fill=tk.X)

        self.log_box = tk.Text(
            pbox, height=7, state='disabled',
            font=("Consolas", 9),
            bg=self._LOG_BG, fg=self._LOG_FG,
            relief=tk.FLAT, bd=0,
        )
        self.log_box.pack(fill=tk.X, pady=(6, 0))

        # Run button
        self.run_btn = tk.Button(
            self.root, text="▶   RUN TRANSFER",
            bg=self._GREEN, fg="white",
            font=("Segoe UI", 13, "bold"),
            activebackground="#229954",
            width=24, height=2,
            relief=tk.FLAT, cursor="hand2",
            command=self._start,
        )
        self.run_btn.pack(pady=14)

    def _file_row(self, parent, label, var, cmd, row):
        tk.Label(
            parent, text=label,
            font=("Segoe UI", 9), bg=self._BG,
            anchor='w', width=30,
        ).grid(row=row, column=0, sticky='w', pady=6)
        tk.Entry(
            parent, textvariable=var, width=40,
            state='readonly', relief=tk.GROOVE,
        ).grid(row=row, column=1, padx=8)
        tk.Button(
            parent, text="Browse …", command=cmd,
            font=("Segoe UI", 8), width=9,
        ).grid(row=row, column=2)

    # ── File browsing ─────────────────────────────────────────────────────────

    def _browse_source(self):
        p = filedialog.askopenfilename(
            title="Select Source PEA Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if p:
            self.source_var.set(p)

    def _browse_template(self):
        p = filedialog.askopenfilename(
            title="Select Arabic Question Bank Template",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if p:
            self.template_var.set(p)

    def _browse_output(self):
        p = filedialog.asksaveasfilename(
            title="Save Output As",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            initialfile="Final_Question_Bank.docx",
        )
        if p:
            self.output_var.set(p)

    # ── Processing ────────────────────────────────────────────────────────────

    def _append_log(self, msg: str):
        self.log_box.config(state='normal')
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state='disabled')

    def _start(self):
        src  = self.source_var.get()
        tmpl = self.template_var.get()
        out  = self.output_var.get()

        if not all([src, tmpl, out]):
            messagebox.showwarning(
                "Missing input",
                "Please select all three files before running.",
            )
            return

        # Clear previous log
        self.log_box.config(state='normal')
        self.log_box.delete('1.0', tk.END)
        self.log_box.config(state='disabled')

        self.run_btn.config(
            text="⏳  Processing …",
            state=tk.DISABLED,
            bg=self._GREY,
        )
        self.progress_bar.start(10)

        threading.Thread(
            target=self._backend,
            args=(src, tmpl, out),
            daemon=True,
        ).start()

    def _backend(self, src, tmpl, out):
        def log(msg):
            self.root.after(0, self._append_log, msg)

        try:
            process_and_transfer(src, tmpl, out, log=log)
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Success  ✓",
                    f"All questions transferred successfully!\n\nSaved to:\n{out}",
                ),
            )
        except Exception as exc:
            err = str(exc)
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Error",
                    f"Transfer failed:\n\n{err}",
                ),
            )
        finally:
            self.root.after(0, self._done)

    def _done(self):
        self.progress_bar.stop()
        self.run_btn.config(
            text="▶   RUN TRANSFER",
            state=tk.NORMAL,
            bg=self._GREEN,
        )


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not _TKINTER_AVAILABLE:
        print("tkinter is not available in this environment.")
        print("Use process_and_transfer() directly from Python instead.")
    else:
        root = tk.Tk()
        app  = TransferApp(root)
        root.mainloop()
