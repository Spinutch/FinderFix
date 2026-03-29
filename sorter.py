#!/usr/bin/env python3
"""
════════════════════════════════════════════════════════════════
  ClassSort AI v2 — Semantic File Organizer + Bulk Review GUI
  macOS | Python 3.8+ | tkinter (built-in) + sentence-transformers
════════════════════════════════════════════════════════════════

FLOW:
  Phase 1 — Background AI scan (progress shown on a loading screen)
  Phase 2 — Tkinter Bulk Review GUI  (treeview, editable destinations)
  Phase 3 — Execute Moves button     (moves/simulates + always writes reports)

KEY CHANGES FROM v1:
  • No dynamic Unsorted subdirectories — unmatched files go directly to
    a single flat  All Classes/Unsorted/  folder.
  • Bulk Review GUI replaces CLI interaction.  Every file is shown in a
    treeview table; click any Destination cell to reassign it.
  • Reports (TXT + CSV) are ALWAYS written to disk — even in DRY_RUN mode.
  • A scrollable Results window opens automatically after execution.

REQUIRED PACKAGES:
    pip install sentence-transformers keybert PyPDF2 python-docx python-pptx openpyxl
"""

# ════════════════════════════════════════════════════════════════
#  SECTION 1 — CONFIGURATION  (edit freely)
# ════════════════════════════════════════════════════════════════

SOURCE_FOLDER  = "/Users/justinevaldes/Desktop/toSort"
CLASSES_FOLDER = "/Users/justinevaldes/Desktop/school"

# Set True  → simulate without moving files.
# Reports are ALWAYS written to disk regardless of this flag.
DRY_RUN = True

# ── Similarity thresholds (cosine similarity: 0.0–1.0) ───────
MIN_SIMILARITY   = 0.28   # below this → Unsorted
STRONG_THRESHOLD = 0.45   # at or above → confident match
AMBIGUITY_GAP    = 0.06   # top-2 classes within this gap → ambiguous

# ── Anchor enrichment ────────────────────────────────────────
# Sample existing files inside each class folder to enrich its
# semantic representation.  Set MAX_ANCHOR_FILES = 0 to disable.
MAX_ANCHOR_FILES = 5
MAX_ANCHOR_CHARS = 2000

# ── Model ─────────────────────────────────────────────────────
# all-MiniLM-L6-v2  → ~80 MB one-time download, fast (~0.3 s/file)
# all-mpnet-base-v2 → ~420 MB, higher accuracy, slower
EMBEDDING_MODEL = "all-MiniLM-L6-v2"

# ── Report filenames ──────────────────────────────────────────
REPORT_TXT = "sorting_report.txt"
REPORT_CSV = "sorting_report.csv"

# ── File type sets ────────────────────────────────────────────
IMAGE_EXTENSIONS   = {".jpg", ".jpeg", ".png", ".gif", ".bmp",
                      ".tiff", ".heic", ".webp", ".svg"}
ARCHIVE_EXTENSIONS = {".zip", ".tar", ".gz", ".rar", ".7z", ".dmg", ".pkg"}


# ════════════════════════════════════════════════════════════════
#  SECTION 2 — IMPORTS
# ════════════════════════════════════════════════════════════════

import os, re, sys, csv, shutil, queue, logging, threading, traceback
from pathlib import Path
from datetime import datetime
from collections import defaultdict

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# ── Optional AI / extraction libraries ───────────────────────
try:
    from sentence_transformers import SentenceTransformer, util as st_util
    HAS_SBERT = True
except ImportError:
    HAS_SBERT = False

try:
    from keybert import KeyBERT          # noqa: F401 (imported for future use)
    HAS_KEYBERT = True
except ImportError:
    HAS_KEYBERT = False

try:
    import PyPDF2;                        HAS_PYPDF2   = True
except ImportError:                       HAS_PYPDF2   = False
try:
    from docx import Document as DocxDoc; HAS_DOCX     = True
except ImportError:                       HAS_DOCX     = False
try:
    from pptx import Presentation as PptxPrs; HAS_PPTX = True
except ImportError:                       HAS_PPTX     = False
try:
    import openpyxl;                      HAS_OPENPYXL = True
except ImportError:                       HAS_OPENPYXL = False

logging.basicConfig(level=logging.INFO, format="%(levelname)-8s %(message)s")
log = logging.getLogger("ClassSort-AI")


# ════════════════════════════════════════════════════════════════
#  SECTION 3 — TEXT EXTRACTION
# ════════════════════════════════════════════════════════════════

def _read_txt(p: Path) -> str:
    try:    return p.read_text(encoding="utf-8", errors="ignore")
    except: return ""

def _read_pdf(p: Path) -> str:
    if not HAS_PYPDF2: return ""
    try:
        parts = []
        with open(p, "rb") as fh:
            for page in PyPDF2.PdfReader(fh).pages:
                parts.append(page.extract_text() or "")
        return " ".join(parts)
    except: return ""

def _read_docx(p: Path) -> str:
    if not HAS_DOCX: return ""
    try:    return " ".join(par.text for par in DocxDoc(str(p)).paragraphs)
    except: return ""

def _read_pptx(p: Path) -> str:
    if not HAS_PPTX: return ""
    try:
        parts = []
        for slide in PptxPrs(str(p)).slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"): parts.append(shape.text)
        return " ".join(parts)
    except: return ""

def _read_xlsx(p: Path) -> str:
    if not HAS_OPENPYXL: return ""
    try:
        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        return " ".join(str(c) for row in wb.active.iter_rows(values_only=True)
                        for c in row if c is not None)
    except: return ""

_EXTRACTORS = {
    ".txt": _read_txt, ".md": _read_txt, ".rtf": _read_txt, ".csv": _read_txt,
    ".pdf": _read_pdf, ".docx": _read_docx,
    ".pptx": _read_pptx, ".xlsx": _read_xlsx,
}

def extract_text(filepath: Path, max_chars: int = 8000) -> str:
    """Extract text from a file and truncate to max_chars."""
    fn = _EXTRACTORS.get(filepath.suffix.lower())
    return (fn(filepath) if fn else "")[:max_chars]


# ════════════════════════════════════════════════════════════════
#  SECTION 4 — SEMANTIC CLASSIFIER
# ════════════════════════════════════════════════════════════════

# Abbreviation map used to expand terse folder names into richer text.
_ABBREV = {
    "CS":   "computer science programming software algorithms",
    "MATH": "mathematics calculus algebra statistics",
    "HIST": "history civilization society politics",
    "ENG":  "english literature writing composition",
    "PHYS": "physics mechanics energy wave",
    "CHEM": "chemistry elements reactions lab",
    "BIO":  "biology cells organisms genetics",
    "ECON": "economics market finance trade",
    "PSYC": "psychology behavior mind cognitive",
    "SOC":  "sociology society culture",
    "STAT": "statistics data probability analysis",
    "NSEC": "network security cybersecurity",
    "NSSE": "network security forensics digital investigation",
    "IT":   "information technology systems",
    "ART":  "art design visual creative",
    "GEO":  "geography spatial environment",
}

def _expand_folder_name(name: str) -> str:
    """
    Turn a terse folder name into a semantically rich description.
    e.g. "CS101" → "CS 101 computer science programming software algorithms"
    """
    m = re.match(r"^([A-Za-z]+)(\d*)(.*)$", name)
    if m:
        prefix = m.group(1).upper()
        rest   = (m.group(2) + " " + m.group(3)).strip()
        expand = _ABBREV.get(prefix, prefix.lower())
        return f"{prefix} {rest} {expand}".strip()
    return re.sub(r"[_\-]+", " ", name)

def _build_anchor(class_dir: Path) -> str:
    """
    Build the full semantic anchor for a class folder:
    expanded name  +  text sampled from existing files in that folder.
    """
    parts = [_expand_folder_name(class_dir.name)]
    if MAX_ANCHOR_FILES > 0 and class_dir.exists():
        sampled = 0
        for f in class_dir.iterdir():
            if sampled >= MAX_ANCHOR_FILES: break
            if not f.is_file() or f.name.startswith("."): continue
            snippet = extract_text(f, MAX_ANCHOR_CHARS)
            if snippet.strip():
                parts.append(snippet)
                sampled += 1
    return " ".join(parts)


class SemanticClassifier:
    """
    Loads the SentenceTransformer model once, discovers class folders
    from the filesystem, and classifies files by cosine similarity.
    """

    def __init__(self, classes_dir: Path, progress_cb=None):
        if not HAS_SBERT:
            raise RuntimeError(
                "sentence-transformers is not installed.\n"
                "Run:  pip install sentence-transformers"
            )

        if progress_cb: progress_cb("Loading AI model  (first run: ~80 MB download)…")
        self.model = SentenceTransformer(EMBEDDING_MODEL)

        # ── Discover class folders dynamically ───────────────
        self.class_dirs = {
            d.name: d for d in classes_dir.iterdir()
            if d.is_dir() and not d.name.startswith(".")
               and d.name not in {"Review", "Unsorted"}
        }
        if not self.class_dirs:
            raise RuntimeError(
                f"No class subfolders found in:\n{classes_dir}\n\n"
                "Please create at least one class folder first."
            )

        # ── Build + encode semantic anchors for all classes ───
        if progress_cb: progress_cb("Building semantic class anchors…")
        self.class_names = sorted(self.class_dirs)
        anchors = [_build_anchor(self.class_dirs[n]) for n in self.class_names]
        self.class_embeddings = self.model.encode(
            anchors, convert_to_tensor=True, show_progress_bar=False
        )

    def classify(self, filepath: Path) -> dict:
        """
        Returns a dict:
            action      : "move" | "review" | "unsorted"
            class_name  : str | None
            similarity  : float (0.0–1.0)
            confidence  : "STRONG" | "LOW" | "AMBIGUOUS" | "NONE"
            basis       : human-readable explanation
        """
        ext = filepath.suffix.lower()

        # Extension shortcuts — no meaningful text to embed
        if ext in IMAGE_EXTENSIONS:
            return self._no_match(0.0, "Image file — no text to embed")
        if ext in ARCHIVE_EXTENSIONS:
            return self._no_match(0.0, "Archive file — no text to embed")

        # Build corpus: filename hint + extracted content
        corpus = (
            filepath.stem.replace("_", " ").replace("-", " ") + " "
            + extract_text(filepath)
        ).strip()

        if not corpus:
            return self._no_match(0.0, "Empty / unreadable — no text extracted")

        # Encode and compare
        file_emb = self.model.encode(corpus, convert_to_tensor=True)
        sims     = st_util.cos_sim(file_emb, self.class_embeddings)[0]

        scores = sorted(
            [(self.class_names[i], float(sims[i])) for i in range(len(self.class_names))],
            key=lambda x: x[1], reverse=True,
        )
        best,  best_score  = scores[0]
        sec_name, sec_score = scores[1] if len(scores) > 1 else ("—", 0.0)

        # Below minimum threshold → no match
        if best_score < MIN_SIMILARITY:
            return self._no_match(
                best_score,
                f"Best similarity {best_score:.3f} below threshold {MIN_SIMILARITY} "
                f"(closest class: {best})",
            )

        # Top two classes too close → ambiguous
        if (best_score - sec_score) <= AMBIGUITY_GAP and sec_score >= MIN_SIMILARITY:
            return {
                "action":     "review",
                "class_name": best,         # offer the best guess as default dest
                "similarity": best_score,
                "confidence": "AMBIGUOUS",
                "basis":      (f"Tied: {best} ({best_score:.3f}) vs "
                               f"{sec_name} ({sec_score:.3f}) — "
                               f"gap {best_score-sec_score:.3f} ≤ {AMBIGUITY_GAP}"),
            }

        # Clear single winner
        confidence = "STRONG" if best_score >= STRONG_THRESHOLD else "LOW"
        return {
            "action":     "move",
            "class_name": best,
            "similarity": best_score,
            "confidence": confidence,
            "basis":      (f"{confidence}: {best} (sim={best_score:.3f}, "
                           f"runner-up={sec_name} @ {sec_score:.3f})"),
        }

    @staticmethod
    def _no_match(sim, reason):
        return {"action": "unsorted", "class_name": None,
                "similarity": sim, "confidence": "NONE", "basis": reason}


# ════════════════════════════════════════════════════════════════
#  SECTION 5 — FILE SYSTEM HELPERS
# ════════════════════════════════════════════════════════════════

def safe_resolve(p: str) -> Path:
    return Path(os.path.expanduser(p)).resolve()

def is_system(f: Path) -> bool:
    return f.name.startswith(".") or f.name in {"Thumbs.db", "desktop.ini"}

def unique_dest(path: Path) -> Path:
    """Return path unchanged, or path_1, path_2 … until a free name is found."""
    if not path.exists(): return path
    i = 1
    while True:
        c = path.parent / f"{path.stem}_{i}{path.suffix}"
        if not c.exists(): return c
        i += 1


# ════════════════════════════════════════════════════════════════
#  SECTION 6 — REPORT GENERATION
#  Always writes TXT + CSV regardless of DRY_RUN.
# ════════════════════════════════════════════════════════════════

def generate_reports(results: list, classes_dir: Path, dry_run: bool) -> str:
    """
    Writes sorting_report.txt and sorting_report.csv.
    Returns the full text of the report for display in the GUI.
    """
    ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mode = "DRY RUN — simulated, no files were moved" if dry_run else "LIVE RUN"

    stats = defaultdict(int)
    for r in results:
        stats[r["action"]] += 1
        stats["total"]     += 1

    sep  = "═" * 72
    dash = "─" * 72
    lines = [
        sep,
        "  CLASSORT AI v2 — SORTING REPORT",
        f"  Generated : {ts}",
        f"  Mode      : {mode}",
        f"  Source    : {SOURCE_FOLDER}",
        f"  Classes   : {CLASSES_FOLDER}",
        f"  Model     : {EMBEDDING_MODEL}",
        sep,
        f"  Total files processed : {stats['total']}",
        f"  Moved / simulated     : {stats.get('moved', 0) + stats.get('simulated', 0)}",
        f"  Errors                : {stats.get('error', 0)}",
        "",
        dash,
        "  PER-FILE DETAILS",
        dash,
    ]
    for r in results:
        lines += [
            f"\n  File        : {r['filename']}",
            f"  Original    : {r['original_path']}",
            f"  Destination : {r['final_dest']}",
            f"  Final path  : {r['final_path']}",
            f"  Action      : {r['action'].upper()}  [{r['confidence']}]",
            f"  Similarity  : {r['similarity']:.4f}",
            f"  Basis       : {r['basis']}",
        ]
    lines.append("\n" + sep)
    report_text = "\n".join(lines)

    txt_path = classes_dir / REPORT_TXT
    csv_path = classes_dir / REPORT_CSV

    # Always write — even in dry-run mode
    try:
        txt_path.write_text(report_text, encoding="utf-8")
        log.info(f"Report → {txt_path}")
    except Exception as e:
        log.warning(f"TXT report write failed: {e}")

    fields = ["filename", "original_path", "final_dest", "final_path",
              "action", "confidence", "similarity", "basis"]
    try:
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows([{k: r.get(k, "") for k in fields} for r in results])
        log.info(f"CSV    → {csv_path}")
    except Exception as e:
        log.warning(f"CSV report write failed: {e}")

    return report_text


# ════════════════════════════════════════════════════════════════
#  SECTION 7 — COLOUR PALETTE & STYLE CONSTANTS
# ════════════════════════════════════════════════════════════════

C_STRONG    = "#E8F5E9"   # soft green  — confident match
C_LOW       = "#FFF9C4"   # soft yellow — low confidence
C_AMBIGUOUS = "#FFE0B2"   # soft orange — ambiguous
C_UNSORTED  = "#FFCDD2"   # soft red    — no match
C_DONE      = "#E3F2FD"   # light blue  — executed/simulated
C_BG        = "#F5F5F5"   # window background
C_HEADER    = "#263238"   # dark header bar
C_ACCENT    = "#1565C0"   # button blue

STATUS_LABELS = {
    "STRONG":    "● Strong Match",
    "LOW":       "◐ Low Confidence",
    "AMBIGUOUS": "⚡ Ambiguous",
    "NONE":      "✗ No Match",
}


# ════════════════════════════════════════════════════════════════
#  SECTION 8 — TKINTER APPLICATION
# ════════════════════════════════════════════════════════════════

class ClassSortApp:
    """
    Main tkinter application.

    Internal state machine:
      loading  → background thread is scanning / loading model
      review   → treeview displayed, user can edit destinations
      done     → execution finished, results window open
    """

    def __init__(self):
        self.source_dir  = safe_resolve(SOURCE_FOLDER)
        self.classes_dir = safe_resolve(CLASSES_FOLDER)
        self.q           = queue.Queue()   # background thread → main thread
        self.files_data  = []              # list of classification dicts
        self.iid_map     = {}              # treeview row iid → files_data index
        self.all_dests   = []              # destination choices for combobox
        self._editor     = None            # currently open inline combobox

        # ── Root window ───────────────────────────────────────
        self.root = tk.Tk()
        self.root.title("ClassSort AI — Bulk File Organizer")
        self.root.geometry("1020x660")
        self.root.minsize(840, 500)
        self.root.configure(bg=C_BG)

        # "clam" theme is required for treeview row background colors on macOS.
        # The native "aqua" theme ignores tag backgrounds.
        style = ttk.Style(self.root)
        style.theme_use("clam")
        self._configure_styles(style)

        self._build_loading_frame()
        self._start_scan()
        self.root.after(150, self._poll_queue)
        self.root.mainloop()

    # ── ttk style configuration ───────────────────────────────
    def _configure_styles(self, s: ttk.Style):
        s.configure("TFrame",        background=C_BG)
        s.configure("Header.TFrame", background=C_HEADER)
        s.configure("Header.TLabel", background=C_HEADER, foreground="white",
                                     font=("Helvetica Neue", 13, "bold"))
        s.configure("Sub.TLabel",    background=C_HEADER, foreground="#B0BEC5",
                                     font=("Helvetica Neue", 10))
        s.configure("Info.TLabel",   background=C_BG,
                                     font=("Helvetica Neue", 11))
        s.configure("Accent.TButton", font=("Helvetica Neue", 12, "bold"),
                                      padding=(18, 8))
        s.configure("Treeview",
                    rowheight=26,
                    font=("Helvetica Neue", 11),
                    background="white",
                    fieldbackground="white",
                    borderwidth=0)
        s.configure("Treeview.Heading",
                    font=("Helvetica Neue", 11, "bold"),
                    background="#ECEFF1", foreground="#263238",
                    relief="flat", padding=(6, 4))
        s.map("Treeview",
              background=[("selected", "#BBDEFB")],
              foreground=[("selected", "#0D47A1")])

    # ════════════════════════════════════════════════════════════
    #  PHASE 1 — LOADING SCREEN
    # ════════════════════════════════════════════════════════════

    def _build_loading_frame(self):
        """Simple centred loading indicator shown while the AI scan runs."""
        self._loading = ttk.Frame(self.root, style="TFrame", padding=50)
        self._loading.place(relx=0.5, rely=0.45, anchor="center")

        tk.Label(self._loading, text="ClassSort AI",
                 font=("Helvetica Neue", 26, "bold"),
                 bg=C_BG, fg=C_ACCENT).pack(pady=(0, 8))

        tk.Label(self._loading,
                 text="Scanning files and building semantic class anchors…",
                 font=("Helvetica Neue", 12), bg=C_BG, fg="#455A64").pack(pady=(0, 24))

        self._pbar = ttk.Progressbar(self._loading, mode="indeterminate", length=420)
        self._pbar.pack(pady=(0, 16))
        self._pbar.start(10)

        self._status_var = tk.StringVar(value="Initialising…")
        tk.Label(self._loading, textvariable=self._status_var,
                 font=("Helvetica Neue", 10), bg=C_BG, fg="#78909C").pack()

    def _set_status(self, text: str):
        """Thread-safe status update for the loading label."""
        self.root.after(0, lambda: self._status_var.set(text))

    # ════════════════════════════════════════════════════════════
    #  BACKGROUND SCAN THREAD
    # ════════════════════════════════════════════════════════════

    def _start_scan(self):
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        """
        Runs entirely in a background thread.
        Communicates with the main thread only via self.q.
        """
        try:
            # Validate paths before loading the heavy model
            for label, path in [("Source", self.source_dir),
                                 ("Classes", self.classes_dir)]:
                if not path.exists():
                    self.q.put({"type": "error",
                                "msg": f"{label} folder not found:\n{path}"})
                    return

            # Initialise classifier (loads model + encodes class anchors)
            classifier  = SemanticClassifier(self.classes_dir, self._set_status)
            class_names = sorted(classifier.class_dirs.keys())

            # Destination options for the combobox: class folders + Unsorted
            dest_options = class_names + ["Unsorted"]

            # Collect source files
            self._set_status("Collecting source files…")
            try:
                all_files = sorted(
                    f for f in self.source_dir.iterdir()
                    if f.is_file() and not is_system(f)
                )
            except PermissionError as e:
                self.q.put({"type": "error", "msg": f"Cannot read source folder:\n{e}"})
                return

            if not all_files:
                self.q.put({"type": "error",
                            "msg": "No files found in the source folder."})
                return

            total   = len(all_files)
            results = []

            for i, filepath in enumerate(all_files, 1):
                self._set_status(f"Classifying {i}/{total}: {filepath.name[:55]}")
                try:
                    r = classifier.classify(filepath)
                except Exception:
                    r = {"action": "unsorted", "class_name": None,
                         "similarity": 0.0, "confidence": "NONE",
                         "basis": f"Error: {traceback.format_exc(limit=1).strip()}"}

                action     = r["action"]
                class_name = r.get("class_name")    # may be set even for "review"
                confidence = r.get("confidence", "NONE")
                similarity = r.get("similarity", 0.0)
                basis      = r.get("basis", "")

                # Determine the AI's suggested destination to show as default.
                # For ambiguous files we still suggest the best-matching class
                # (the user can override it in the GUI).
                if action in ("move", "review") and class_name:
                    default_dest = class_name
                else:
                    default_dest = "Unsorted"

                results.append({
                    "index":        i - 1,
                    "filepath":     filepath,
                    "filename":     filepath.name,
                    "original_path": str(filepath),
                    "ai_dest":      default_dest,   # immutable AI suggestion
                    "current_dest": default_dest,   # user may override in GUI
                    "similarity":   similarity,
                    "confidence":   confidence,
                    "action":       action,
                    "basis":        basis,
                    # filled after Execute:
                    "final_dest":   default_dest,
                    "final_path":   "",
                    "executed":     False,
                })

            self.q.put({
                "type":         "done",
                "results":      results,
                "dest_options": dest_options,
            })

        except Exception as e:
            self.q.put({"type": "error", "msg": str(e)})

    # ── Queue poller ──────────────────────────────────────────
    def _poll_queue(self):
        """Check for background thread results every 150 ms."""
        try:
            msg = self.q.get_nowait()
            if msg["type"] == "done":
                self._pbar.stop()
                self._loading.destroy()
                self.files_data = msg["results"]
                self.all_dests  = msg["dest_options"]
                self._build_review_ui()
            elif msg["type"] == "error":
                self._pbar.stop()
                self._loading.destroy()
                messagebox.showerror("ClassSort — Error", msg["msg"])
                self.root.destroy()
                return
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    # ════════════════════════════════════════════════════════════
    #  PHASE 2 — BULK REVIEW GUI
    # ════════════════════════════════════════════════════════════

    def _build_review_ui(self):
        # ── Header ────────────────────────────────────────────
        hdr = ttk.Frame(self.root, style="Header.TFrame", padding=(16, 11))
        hdr.pack(fill="x")
        ttk.Label(hdr, text="ClassSort AI — Bulk Review",
                  style="Header.TLabel").pack(side="left")
        mode_lbl = "  🔵 DRY RUN  " if DRY_RUN else "  🟢 LIVE  "
        ttk.Label(hdr, text=mode_lbl, style="Sub.TLabel").pack(side="left", padx=8)
        n = len(self.files_data)
        ttk.Label(hdr, text=f"{n} file{'s' if n != 1 else ''} ready for review",
                  style="Sub.TLabel").pack(side="right", padx=6)

        # ── Instruction strip ─────────────────────────────────
        instr = tk.Frame(self.root, bg="#ECEFF1")
        instr.pack(fill="x")
        tk.Label(instr,
                 text=" ↖  Click any cell in the  Destination  column to reassign a file."
                      "  Highlighted rows need your attention.  Then press  Execute Moves.",
                 font=("Helvetica Neue", 10), bg="#ECEFF1", fg="#455A64",
                 pady=5).pack(side="left", padx=10)

        # ── Treeview container ────────────────────────────────
        tv_frame = ttk.Frame(self.root, padding=(10, 6, 10, 0))
        tv_frame.pack(fill="both", expand=True)

        cols = ("filename", "destination", "confidence", "status")
        self.tree = ttk.Treeview(tv_frame, columns=cols,
                                 show="headings", selectmode="browse")

        # Column definitions
        self.tree.heading("filename",    text="Original Filename",  anchor="w")
        self.tree.heading("destination", text="▼ Destination",      anchor="w")
        self.tree.heading("confidence",  text="Confidence",         anchor="center")
        self.tree.heading("status",      text="Status",             anchor="w")

        self.tree.column("filename",    width=330, minwidth=180, stretch=True,  anchor="w")
        self.tree.column("destination", width=190, minwidth=120, stretch=False, anchor="w")
        self.tree.column("confidence",  width=100, minwidth=80,  stretch=False, anchor="center")
        self.tree.column("status",      width=175, minwidth=130, stretch=False, anchor="w")

        # Row colour tags — requires "clam" theme (set in __init__)
        self.tree.tag_configure("STRONG",    background=C_STRONG)
        self.tree.tag_configure("LOW",       background=C_LOW)
        self.tree.tag_configure("AMBIGUOUS", background=C_AMBIGUOUS)
        self.tree.tag_configure("NONE",      background=C_UNSORTED)
        self.tree.tag_configure("DONE",      background=C_DONE)

        # Scrollbars
        vsb = ttk.Scrollbar(tv_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tv_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        # Bind click for inline destination editing
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

        # ── Legend ────────────────────────────────────────────
        legend = tk.Frame(self.root, bg=C_BG)
        legend.pack(fill="x", padx=12, pady=(4, 0))
        tk.Label(legend, text="Legend: ", font=("Helvetica Neue", 9, "bold"),
                 bg=C_BG).pack(side="left")
        for colour, label in [
            (C_STRONG,    "Strong Match"),
            (C_LOW,       "Low Confidence"),
            (C_AMBIGUOUS, "Ambiguous — needs review"),
            (C_UNSORTED,  "No Match → Unsorted"),
        ]:
            tk.Label(legend, text=f"  {label}  ", bg=colour,
                     font=("Helvetica Neue", 9), padx=5, pady=2,
                     relief="flat").pack(side="left", padx=(0, 8))

        # ── Bottom action bar ─────────────────────────────────
        bottom = tk.Frame(self.root, bg=C_BG)
        bottom.pack(fill="x", padx=12, pady=8)

        self._stats_var = tk.StringVar()
        tk.Label(bottom, textvariable=self._stats_var,
                 font=("Helvetica Neue", 10), bg=C_BG, fg="#607D8B").pack(side="left")

        btn_label = "  🔵 Simulate Moves (Dry Run)  " if DRY_RUN else "  ✅ Execute Moves  "
        self._exec_btn = ttk.Button(
            bottom, text=btn_label, style="Accent.TButton",
            command=self._execute_moves,
        )
        self._exec_btn.pack(side="right")

        # Populate rows and stats
        self._populate_tree()
        self._refresh_stats()

    # ── Populate treeview ─────────────────────────────────────
    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.iid_map.clear()

        for data in self.files_data:
            idx  = data["index"]
            iid  = str(idx)
            conf = data["confidence"]
            pct  = f"{data['similarity'] * 100:.1f}%"
            self.tree.insert(
                "", "end", iid=iid,
                values=(data["filename"], data["current_dest"],
                        pct, STATUS_LABELS.get(conf, conf)),
                tags=(conf,),
            )
            self.iid_map[iid] = idx

    def _refresh_stats(self):
        c = defaultdict(int)
        for d in self.files_data: c[d["confidence"]] += 1
        self._stats_var.set(
            f"Total: {len(self.files_data)}    "
            f"Strong: {c['STRONG']}    "
            f"Low: {c['LOW']}    "
            f"Ambiguous: {c['AMBIGUOUS']}    "
            f"Unmatched: {c['NONE']}"
        )

    # ════════════════════════════════════════════════════════════
    #  INLINE DESTINATION EDITOR
    # ════════════════════════════════════════════════════════════

    def _dismiss_editor(self):
        """Safely destroy the floating combobox if it exists."""
        if self._editor and self._editor.winfo_exists():
            self._editor.destroy()
        self._editor = None

    def _on_tree_click(self, event):
        """
        Fires on every mouse release inside the treeview.
        Opens a combobox editor only when the click lands on the
        destination column of a non-executed row.
        """
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            self._dismiss_editor()
            return

        col = self.tree.identify_column(event.x)  # "#1", "#2", …
        iid = self.tree.identify_row(event.y)

        if col != "#2" or not iid:                 # only destination column
            self._dismiss_editor()
            return

        idx = self.iid_map.get(iid)
        if idx is None or self.files_data[idx].get("executed"):
            return

        # Small delay ensures any prior dismiss fires first
        self.root.after(30, lambda i=iid: self._open_dest_editor(i))

    def _open_dest_editor(self, iid: str):
        """
        Place a ttk.Combobox directly over the destination cell.
        Uses the cell's bounding box (relative to the treeview widget)
        for precise positioning.
        """
        self._dismiss_editor()

        bbox = self.tree.bbox(iid, "destination")
        if not bbox:
            return   # row scrolled out of view

        x, y, w, h = bbox
        idx         = self.iid_map[iid]
        current     = self.files_data[idx]["current_dest"]

        cb = ttk.Combobox(self.tree, values=self.all_dests,
                          state="readonly", font=("Helvetica Neue", 11))
        cb.set(current)
        cb.place(x=x, y=y, width=w, height=h)
        cb.focus_set()
        self._editor = cb

        def on_select(_event):
            new_dest = cb.get()
            self.files_data[idx]["current_dest"] = new_dest
            old = list(self.tree.item(iid, "values"))
            old[1] = new_dest
            self.tree.item(iid, values=old)
            self._dismiss_editor()
            self._refresh_stats()

        cb.bind("<<ComboboxSelected>>", on_select)
        cb.bind("<FocusOut>",            lambda e: self._dismiss_editor())
        cb.bind("<Escape>",              lambda e: self._dismiss_editor())

    # ════════════════════════════════════════════════════════════
    #  PHASE 3 — EXECUTE MOVES
    # ════════════════════════════════════════════════════════════

    def _execute_moves(self):
        self._dismiss_editor()

        n    = len(self.files_data)
        mode = "DRY RUN (simulation)" if DRY_RUN else "LIVE"
        note = ("Files will NOT be moved — this is a simulation.\n"
                if DRY_RUN else
                "Files will be PERMANENTLY MOVED to their destinations.\n")

        if not messagebox.askyesno(
            "Confirm",
            f"Mode: {mode}\n\n"
            f"{note}"
            f"Processing {n} file{'s' if n != 1 else ''}.\n\n"
            "Reports (TXT + CSV) will be written to the Classes folder either way.\n\n"
            "Proceed?",
        ):
            return

        self._exec_btn.configure(state="disabled", text="  Working…  ")
        self.root.update()

        unsorted_dir = self.classes_dir / "Unsorted"
        report_rows  = []

        for data in self.files_data:
            filepath  = data["filepath"]
            dest_name = data["current_dest"]   # user's final choice
            iid       = str(data["index"])

            # Resolve the physical destination directory
            if dest_name == "Unsorted":
                dest_dir = unsorted_dir
            else:
                dest_dir = self.classes_dir / dest_name

            # Build a safe (non-overwriting) destination path
            raw_path  = dest_dir / filepath.name
            safe_path = unique_dest(raw_path)

            if DRY_RUN:
                action_taken = "SIMULATED"
                final_path   = str(safe_path)
            else:
                dest_dir.mkdir(parents=True, exist_ok=True)
                try:
                    shutil.move(str(filepath), str(safe_path))
                    action_taken = "MOVED"
                    final_path   = str(safe_path)
                except Exception as e:
                    action_taken = f"ERROR"
                    final_path   = f"(failed) {e}"
                    log.error(f"Move failed for {filepath.name}: {e}")

            # Update data record
            data["final_dest"]  = dest_name
            data["final_path"]  = final_path
            data["action"]      = action_taken.lower()
            data["executed"]    = True

            # Update treeview row: change status column + colour to DONE
            done_label = ("✓ Moved"      if action_taken == "MOVED"      else
                          "~ Simulated"  if action_taken == "SIMULATED"  else
                          "✗ Error")
            vals = list(self.tree.item(iid, "values"))
            vals[3] = done_label
            self.tree.item(iid, values=vals, tags=("DONE",))

            report_rows.append({
                "filename":     data["filename"],
                "original_path": data["original_path"],
                "final_dest":   dest_name,
                "final_path":   final_path,
                "action":       action_taken,
                "confidence":   data["confidence"],
                "similarity":   data["similarity"],
                "basis":        data["basis"],
            })

            # Flush UI every 10 files to stay responsive
            if data["index"] % 10 == 0:
                self.root.update()

        # Generate + save reports (always)
        report_text = generate_reports(report_rows, self.classes_dir, DRY_RUN)

        self._exec_btn.configure(state="normal", text="  ✓ Done  ")
        self._show_results_window(report_text)

    # ════════════════════════════════════════════════════════════
    #  RESULTS WINDOW — always shown, even in dry-run mode
    # ════════════════════════════════════════════════════════════

    def _show_results_window(self, report_text: str):
        win = tk.Toplevel(self.root)
        win.title("ClassSort AI — Sorting Report")
        win.geometry("860x580")
        win.minsize(600, 400)
        win.configure(bg=C_BG)

        # Header
        hdr = tk.Frame(win, bg=C_HEADER)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  Sorting Report",
                 font=("Helvetica Neue", 13, "bold"),
                 bg=C_HEADER, fg="white", pady=10).pack(side="left")
        mode_tag = "  🔵 DRY RUN — no files moved  " if DRY_RUN else "  🟢 LIVE — files moved  "
        tk.Label(hdr, text=mode_tag,
                 font=("Helvetica Neue", 10),
                 bg=C_HEADER, fg="#B0BEC5").pack(side="right", padx=10)

        # Scrollable text area showing the full report
        txt = scrolledtext.ScrolledText(
            win, wrap="none",
            font=("Menlo", 10),
            bg="#FAFAFA", relief="flat",
            padx=14, pady=14,
        )
        txt.pack(fill="both", expand=True, padx=10, pady=(10, 4))
        txt.insert("1.0", report_text)
        txt.configure(state="disabled")

        # Footer showing where the files were saved
        footer_txt = (
            f"📄 {self.classes_dir / REPORT_TXT}\n"
            f"📊 {self.classes_dir / REPORT_CSV}"
        )
        tk.Label(win, text=footer_txt,
                 font=("Helvetica Neue", 9), bg=C_BG, fg="#78909C",
                 justify="left").pack(anchor="w", padx=14, pady=(0, 4))

        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 10))


# ════════════════════════════════════════════════════════════════
#  SECTION 9 — ENTRY POINT
# ════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if not HAS_SBERT:
        print("\n[ERROR] sentence-transformers is not installed.")
        print("Run:  pip install sentence-transformers keybert")
        print("      pip install PyPDF2 python-docx python-pptx openpyxl\n")
        sys.exit(1)

    ClassSortApp()