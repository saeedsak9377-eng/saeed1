#!/usr/bin/env python3
"""
mcq_transfer.py — MCQ Bank Transfer Tool
=========================================
Transfers every MCQ question from a PEA-format source document into a
paragraph-structured Arabic question-bank output document.

Source document format (PEA / EXAMKEY)
---------------------------------------
Table 0 (EXAMKEY) columns:
    0  Seq          – sequential number (1, 2, 3 …)
    1  Question ID  – internal reference  (e.g. "578621 (3)")
    2  External ID  – formal question code (e.g. "TSQ10084-01") ← رمز السؤال
    3  Blueprint
    4  Answer       – correct option  (A / B / C / D)
    5  Status

Body paragraphs following the EXAMKEY table:
    "QUESTION N"
    <question body lines>
    "A. <option text>"
    "B. <option text>"
    "C. <option text>"
    "D. <option text>"

Output document format (mirrors the Arabic question-bank template)
------------------------------------------------------------------
    سري للغاية | Top Secret
    سؤال رقم: N
    رمز السؤال : TSQ10084-01
    نص السؤال :
    <question body paragraphs – text / math / images preserved>
    الجواب الصحيح (*)
    نص الاجابة\t\tرقم الاجابة
    [* ]<option A text>\t\tأ
    [* ]<option B text>\t\tب
    [* ]<option C text>\t\tت
    [* ]<option D text>\t\tث
    ── page break ──

The star prefix  *  appears only on the correct-answer option.

Usage
-----
    python mcq_transfer.py        # opens the GUI
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

# English option letter → Arabic option letter
OPTION_MAP: dict = {'A': 'أ', 'B': 'ب', 'C': 'ت', 'D': 'ث'}
OPTION_MAP_REV: dict = {v: k for k, v in OPTION_MAP.items()}

# XML namespace URIs
NS_MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS_WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_REL  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMBED_ATTR = f'{{{NS_REL}}}embed'

# Boilerplate lines in the source document that should be silently skipped
SKIP_PHRASES = (
    "REFER TO THE FOLLOWING",
    "أسئلة الاختيار",
    "فيما يلي سؤال",
    "أسئلة المقارنة",
    "EXHIBIT",
    "% of times",           # distribution statistics block in the header
)

# EXAMKEY table column indices
COL_SEQ     = 0
COL_EXT_ID  = 2   # External ID → goes into  رمز السؤال
COL_ANSWER  = 4   # Correct answer letter (A / B / C / D)


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 1 – PARSE SOURCE DOCUMENT
# ─────────────────────────────────────────────────────────────────────────────

def extract_exam_key(doc) -> dict:
    """
    Read the EXAMKEY table (first table in the source document).

    Returns
    -------
    dict  {question_number (int) →
               {'id': External_ID (str), 'answer': Arabic_letter (str)}}
    """
    if not doc.tables:
        raise ValueError(
            "Source document has no tables.\n"
            "Expected the EXAMKEY as the very first table with columns:\n"
            "  Seq | Question ID | External ID | Blueprint | Answer | Status"
        )

    key_data: dict = {}
    for row in doc.tables[0].rows[1:]:   # row 0 = header — skip
        cells = row.cells
        if len(cells) <= max(COL_SEQ, COL_EXT_ID, COL_ANSWER):
            continue
        seq     = cells[COL_SEQ].text.strip()
        ext_id  = cells[COL_EXT_ID].text.strip()
        ans_eng = cells[COL_ANSWER].text.strip().upper()
        if seq.isdigit():
            key_data[int(seq)] = {
                'id':     ext_id,
                'answer': OPTION_MAP.get(ans_eng, ''),
            }
    return key_data


def _para_has_content(para) -> bool:
    """True when a paragraph carries visible text or an embedded element."""
    if para.text.strip():
        return True
    el = para._element
    return (
        bool(el.xpath('.//m:oMath',   namespaces={'m': NS_MATH})) or
        bool(el.xpath('.//w:drawing', namespaces={'w': NS_WORD}))
    )


def parse_questions(doc) -> dict:
    """
    Walk every paragraph in the source document body and group paragraphs
    into question blocks.

    Returns
    -------
    dict  {question_number (int) →
               {'body': [...], 'A': [...], 'B': [...], 'C': [...], 'D': [...]}}
    """
    questions:       dict = {}
    current_q:       int  = None
    current_section: str  = None   # 'body' | 'A' | 'B' | 'C' | 'D'

    re_question = re.compile(r'^QUESTION\s+(\d+)', re.IGNORECASE)
    re_option   = re.compile(r'^([A-D])\.\s*(.*)', re.DOTALL | re.IGNORECASE)

    for para in doc.paragraphs:
        text = para.text.strip()

        # Skip completely blank paragraphs with no embedded content
        if not text and not _para_has_content(para):
            continue

        # ── New QUESTION header ──────────────────────────────────────────────
        m = re_question.match(text)
        if m:
            current_q       = int(m.group(1))
            current_section = 'body'
            questions[current_q] = {'body': [], 'A': [], 'B': [], 'C': [], 'D': []}
            continue

        # Nothing active yet (header boilerplate before first QUESTION)
        if current_q is None:
            continue

        # Skip boilerplate / instructional lines
        if any(phrase in text for phrase in SKIP_PHRASES):
            continue

        # ── Option header line  (A. / B. / C. / D.) ─────────────────────────
        m = re_option.match(text)
        if m:
            letter = m.group(1).upper()
            current_section = letter
            questions[current_q][letter].append(para)
            continue

        # ── Regular content ───────────────────────────────────────────────────
        if current_section:
            questions[current_q][current_section].append(para)

    return questions


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 2 – XML SURGERY  (preserve text / OMML math / images)
# ─────────────────────────────────────────────────────────────────────────────

def _transfer_image_rels(p_xml, source_doc, target_doc):
    """
    Re-map every  r:embed  relationship inside p_xml from source to target.
    Called after deep-copying a paragraph that may contain embedded images.
    """
    for elem in p_xml.iter():
        if EMBED_ATTR in elem.attrib:
            old_rid = elem.attrib[EMBED_ATTR]
            try:
                rel     = source_doc.part.rels[old_rid]
                new_rid = target_doc.part.relate_to(rel.target_part, rel.reltype)
                elem.attrib[EMBED_ATTR] = new_rid
            except (KeyError, AttributeError):
                pass    # best-effort; image may still render via the original rId


def _append_to_body(doc, p_elem):
    """
    Insert a  <w:p>  element into the document body, just before the
    final  <w:sectPr>  so page-setup properties are preserved.
    """
    body    = doc.element.body
    sect_pr = body.find(qn('w:sectPr'))
    if sect_pr is not None:
        body.insert(list(body).index(sect_pr), p_elem)
    else:
        body.append(p_elem)


def _copy_source_para(doc, src_para, source_doc):
    """
    Deep-copy a source paragraph (with math equations and images intact)
    and append it to doc.
    """
    new_p = deepcopy(src_para._element)
    _transfer_image_rels(new_p, source_doc, doc)
    _append_to_body(doc, new_p)


def _strip_option_letter_from_xml(p_elem):
    """
    Remove the leading  'A. '  /  'B. '  text from the first run of a
    copied option paragraph.  Only touches the very first  <w:t>  node.
    """
    for run in p_elem.findall(qn('w:r')):
        for t in run.findall(qn('w:t')):
            if t.text:
                cleaned = re.sub(r'^[A-D]\.\s*', '', t.text, flags=re.IGNORECASE)
                if cleaned != t.text:
                    t.text = cleaned
                    if cleaned:
                        t.set(qn('xml:space'), 'preserve')
                    else:
                        # Run is now empty — remove it entirely
                        run.getparent().remove(run)
                return   # only process the first text element


def _build_option_para(doc, option_paras, arabic_letter, is_correct, source_doc):
    """
    Build and append one option row paragraph with the format:

        [* ]<option content>\\t\\t<arabic_letter>

    The  *  prefix appears only when  is_correct  is True.
    Option content is copied via deep-XML so math equations and images survive.
    """
    if not option_paras:
        # No content for this option — show the label placeholder
        doc.add_paragraph(f"\t\t{arabic_letter}")
        return

    # Deep-copy the first (and usually only) option paragraph
    new_p = deepcopy(option_paras[0]._element)
    _transfer_image_rels(new_p, source_doc, doc)

    # Strip the  'A. '  letter prefix from the copied text
    _strip_option_letter_from_xml(new_p)

    # Prepend the correct-answer star when needed
    if is_correct:
        star_run = OxmlElement('w:r')
        star_t   = OxmlElement('w:t')
        star_t.text = '* '
        star_t.set(qn('xml:space'), 'preserve')
        star_run.append(star_t)
        pPr = new_p.find(qn('w:pPr'))
        pos = (list(new_p).index(pPr) + 1) if pPr is not None else 0
        new_p.insert(pos, star_run)

    # Append the  \t\tأ  /  \t\tب  …  suffix
    label_run = OxmlElement('w:r')
    label_t   = OxmlElement('w:t')
    label_t.text = f'\t\t{arabic_letter}'
    label_t.set(qn('xml:space'), 'preserve')
    label_run.append(label_t)
    new_p.append(label_run)

    _append_to_body(doc, new_p)

    # Handle rare multi-paragraph options (e.g. a diagram below the text)
    for extra in option_paras[1:]:
        if _para_has_content(extra):
            _copy_source_para(doc, extra, source_doc)


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 3 – BUILD ONE QUESTION BLOCK
# ─────────────────────────────────────────────────────────────────────────────

def build_question_block(doc, q_num, k_data, q_data, source_doc):
    """
    Append one complete question block to doc, exactly matching the
    Arabic question-bank template structure:

        سري للغاية | Top Secret
        سؤال رقم: N
        رمز السؤال : <External ID>
        نص السؤال :
        <body paragraphs>
        الجواب الصحيح (*)
        نص الاجابة\\t\\tرقم الاجابة
        [*]<option A>\\t\\tأ
        [*]<option B>\\t\\tب
        [*]<option C>\\t\\tت
        [*]<option D>\\t\\tث
        ── page break ──
    """
    correct_arabic = k_data.get('answer', '')
    q_id           = k_data.get('id', '')

    # ── Static header lines ───────────────────────────────────────────────────
    doc.add_paragraph("سري للغاية | Top Secret")
    doc.add_paragraph(f"سؤال رقم: {q_num}")
    doc.add_paragraph(f"رمز السؤال : {q_id}")
    doc.add_paragraph("نص السؤال :")

    # ── Question body  (text / math / images preserved via XML copy) ─────────
    body_paras = [p for p in q_data.get('body', []) if _para_has_content(p)]
    if body_paras:
        for p in body_paras:
            _copy_source_para(doc, p, source_doc)
    else:
        doc.add_paragraph("")   # blank placeholder if body is empty

    # ── Answer section ────────────────────────────────────────────────────────
    doc.add_paragraph("الجواب الصحيح (*)")
    doc.add_paragraph("نص الاجابة\t\tرقم الاجابة")

    for eng_letter, arabic_letter in OPTION_MAP.items():
        option_paras = q_data.get(eng_letter, [])
        is_correct   = (arabic_letter == correct_arabic)
        _build_option_para(doc, option_paras, arabic_letter, is_correct, source_doc)

    # ── Page break between questions ──────────────────────────────────────────
    doc.add_page_break()


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 4 – MASTER ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

def _clear_body(doc):
    """
    Remove all paragraphs and tables from the document body so we can
    rebuild the content from scratch, while keeping the  <w:sectPr>  so
    the page layout (size, margins, RTL section direction) is preserved.
    """
    body = doc.element.body
    for elem in list(body):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag != 'sectPr':
            body.remove(elem)
    # Ensure at least one sectPr exists so Word doesn't complain
    if not body.findall(qn('w:sectPr')):
        body.append(OxmlElement('w:sectPr'))


def process_and_transfer(
    source_path:   str,
    template_path: str,
    output_path:   str,
    log=None,
):
    """
    Full pipeline:
        load source → extract key → parse questions →
        init output from template → build blocks → save

    Parameters
    ----------
    source_path   : PEA source .docx
    template_path : Arabic question-bank template .docx
    output_path   : destination .docx
    log           : optional callable(str) for progress messages
    """
    def _log(msg: str):
        if log:
            log(msg)

    _log("Loading source document …")
    source_doc = Document(source_path)

    _log("Extracting exam key …")
    exam_key = extract_exam_key(source_doc)
    if not exam_key:
        raise ValueError(
            "No data found in the EXAMKEY table.\n"
            "Ensure the first table in the source document follows:\n"
            "  Seq | Question ID | External ID | Blueprint | Answer | Status"
        )
    _log(f"  → {len(exam_key)} entries in exam key.")

    _log("Parsing questions …")
    parsed_qs = parse_questions(source_doc)
    if not parsed_qs:
        raise ValueError(
            "No questions found in the source document.\n"
            "Questions must start with 'QUESTION 1', 'QUESTION 2', etc."
        )
    _log(f"  → {len(parsed_qs)} questions parsed.")

    _log("Initialising output document …")
    output_doc = Document(template_path)
    _clear_body(output_doc)

    all_nums = sorted(set(exam_key) & set(parsed_qs))
    if not all_nums:
        raise ValueError(
            "Question numbers in the exam key don't match the parsed questions.\n"
            f"Exam key : {sorted(exam_key.keys())}\n"
            f"Parsed   : {sorted(parsed_qs.keys())}"
        )

    total = len(all_nums)
    _log(f"Writing {total} question block(s) …")

    for idx, q_num in enumerate(all_nums, start=1):
        build_question_block(
            output_doc,
            q_num,
            exam_key[q_num],
            parsed_qs[q_num],
            source_doc,
        )
        _log(
            f"  [{idx}/{total}]  Q{q_num}  →  "
            f"{exam_key[q_num]['id']}  "
            f"(correct: {exam_key[q_num]['answer']})"
        )

    _log(f"Saving → {output_path}")
    output_doc.save(output_path)
    _log("Done!  ✓")


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 5 – GUI
# ─────────────────────────────────────────────────────────────────────────────

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    _TKINTER_AVAILABLE = True
except ImportError:
    _TKINTER_AVAILABLE = False


class TransferApp:
    """Main application window."""

    _BG     = "#ecf0f1"
    _HEADER = "#2c3e50"
    _GREEN  = "#27ae60"
    _GREY   = "#7f8c8d"
    _LOG_BG = "#1a1a2e"
    _LOG_FG = "#00ff88"

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

        self.progress_bar = ttk.Progressbar(pbox, mode='indeterminate', length=580)
        self.progress_bar.pack(fill=tk.X)

        self.log_box = tk.Text(
            pbox, height=7, state='disabled',
            font=("Consolas", 9), bg=self._LOG_BG, fg=self._LOG_FG,
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
            font=("Segoe UI", 9), bg=self._BG, anchor='w', width=30,
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

        self.log_box.config(state='normal')
        self.log_box.delete('1.0', tk.END)
        self.log_box.config(state='disabled')

        self.run_btn.config(text="⏳  Processing …", state=tk.DISABLED, bg=self._GREY)
        self.progress_bar.start(10)

        threading.Thread(
            target=self._backend, args=(src, tmpl, out), daemon=True,
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
                lambda: messagebox.showerror("Error", f"Transfer failed:\n\n{err}"),
            )
        finally:
            self.root.after(0, self._done)

    def _done(self):
        self.progress_bar.stop()
        self.run_btn.config(text="▶   RUN TRANSFER", state=tk.NORMAL, bg=self._GREEN)


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not _TKINTER_AVAILABLE:
        print("tkinter is not available in this environment.")
        print("Import and call  process_and_transfer()  directly from Python.")
    else:
        root = tk.Tk()
        app  = TransferApp(root)
        root.mainloop()
