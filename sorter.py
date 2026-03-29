#!/usr/bin/env python3
"""
════════════════════════════════════════════════════════════════
  ClassSort AI — Semantic Academic File Organizer
  Compatible with macOS | Python 3.8+
════════════════════════════════════════════════════════════════

HOW IT WORKS (AI approach):
  1. Scans your "All Classes" folder to discover class folders dynamically.
     No hardcoded keywords — the folder names themselves are the categories.
  2. Builds a "semantic anchor" for each class folder by combining:
       • The folder name (e.g. "CS101", "MATH203")
       • A short human description (auto-generated from existing files
         found inside the folder, if any)
  3. Extracts text from each source file (PDF, DOCX, PPTX, XLSX, TXT…)
  4. Encodes everything into dense vector embeddings using the
     sentence-transformers model  all-MiniLM-L6-v2  (runs fully
     locally, no API key, ~80 MB one-time download).
  5. Computes cosine similarity between the file embedding and every
     class anchor embedding, then applies configurable thresholds:
       • Strong match  → move to that class folder
       • Ambiguous     → two or more classes score within AMBIGUITY_GAP
                         of each other  → send to Review/
       • No match      → score below MIN_SIMILARITY  → Unsorted/
  6. For Unsorted files, KeyBERT (also local, uses same model) extracts
     the 2 most descriptive keywords from the file's content and uses
     them to name the destination subfolder  (e.g. "Unsorted/Tax_Returns").

WHY sentence-transformers?
  • Runs 100 % locally — no API key, no internet needed after install.
  • Understands MEANING, not just literal keywords:
      "kinematics" → similar to "Physics 150" even without the word "physics".
  • Lightweight (all-MiniLM-L6-v2 is 80 MB, inference < 0.5 s per file).
  • Beginner-friendly: one pip install, two lines of code to use.
  • More robust than fuzzy matching for course-content language.

REQUIRED PACKAGES (see Setup Instructions):
    pip install sentence-transformers keybert PyPDF2 python-docx python-pptx openpyxl

USAGE:
    python file_organizer_ai.py          # dry run by default
    Set DRY_RUN = False below for a live run
"""

# ════════════════════════════════════════════════════════════════
#  SECTION 1 — CONFIGURATION  (edit this section freely)
# ════════════════════════════════════════════════════════════════

# ── Paths ─────────────────────────────────────────────────────
SOURCE_FOLDER  = "/Users/justinevaldes/Desktop/toSort"
CLASSES_FOLDER = "/Users/justinevaldes/Desktop/school"

# ── Safety toggle ─────────────────────────────────────────────
DRY_RUN = True          # ← set to False when you're ready to move files

# ── Similarity thresholds ─────────────────────────────────────
#
#  Cosine similarity ranges from 0.0 (completely unrelated)
#  to 1.0 (identical meaning).  Typical useful range: 0.20–0.70.
#
#  MIN_SIMILARITY   — scores below this value mean "no match at all".
#                     Files go to Unsorted/.   Raise to be stricter.
#
#  STRONG_THRESHOLD — scores at or above this are considered a confident
#                     match.  Below this (but above MIN_SIMILARITY) the
#                     file is still moved but flagged LOW_CONFIDENCE.
#
#  AMBIGUITY_GAP    — if the top-2 class scores are within this distance
#                     of each other, the file is ambiguous → Review/.
#                     Lower = more files go to Review.

MIN_SIMILARITY   = 0.28   # below this → Unsorted
STRONG_THRESHOLD = 0.45   # at or above → confident move
AMBIGUITY_GAP    = 0.06   # if top two classes are this close → Review

# ── Folder anchor enrichment ──────────────────────────────────
#
#  To make sparse class names like "CS101" more meaningful, the script
#  samples text from existing files already inside each class folder.
#  MAX_ANCHOR_FILES controls how many files to sample per class.
#  Set to 0 to disable (use folder name only).

MAX_ANCHOR_FILES = 5       # max existing files to sample per class folder
MAX_ANCHOR_CHARS = 2000    # max characters to read from each anchor file

# ── Unsorted label generation ─────────────────────────────────
#
#  KeyBERT extracts the N most representative keywords from the file.
#  These are joined with underscores to form the subfolder name.
#  E.g. keywords ["tax", "returns"] → "Unsorted/Tax_Returns"

UNSORTED_KEYPHRASE_COUNT  = 2   # how many keywords to extract
UNSORTED_MAX_KEYPHRASE_LEN = 2  # max words per keyphrase (1 or 2 recommended)

# ── Report filenames ──────────────────────────────────────────
REPORT_TXT_FILENAME = "sorting_report.txt"
REPORT_CSV_FILENAME = "sorting_report.csv"

# ── Model name ────────────────────────────────────────────────
#  all-MiniLM-L6-v2 is fast (~0.3s/file) and accurate enough for this task.
#  For higher accuracy at the cost of speed, try: all-mpnet-base-v2
EMBEDDING_MODEL = "all-MiniLM-L6-v2"

# ── File type sets ────────────────────────────────────────────
EXTRACTABLE_EXTENSIONS = {".txt", ".pdf", ".docx", ".pptx", ".xlsx", ".md", ".rtf", ".csv"}
IMAGE_EXTENSIONS       = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".heic", ".webp", ".svg"}
ARCHIVE_EXTENSIONS     = {".zip", ".tar", ".gz", ".rar", ".7z", ".dmg", ".pkg"}

# ════════════════════════════════════════════════════════════════
#  SECTION 2 — IMPORTS
# ════════════════════════════════════════════════════════════════

import os
import re
import sys
import csv
import shutil
import logging
import traceback
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# ── Core AI libraries ─────────────────────────────────────────
try:
    from sentence_transformers import SentenceTransformer, util
    HAS_SBERT = True
except ImportError:
    HAS_SBERT = False

try:
    from keybert import KeyBERT
    HAS_KEYBERT = True
except ImportError:
    HAS_KEYBERT = False

# ── Text extraction libraries ─────────────────────────────────
try:
    import PyPDF2
    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False

try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from pptx import Presentation as PptxPresentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ── Logging ───────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)-8s %(message)s")
log = logging.getLogger("ClassSort-AI")


# ════════════════════════════════════════════════════════════════
#  SECTION 3 — TEXT EXTRACTION
#  Each function returns a plain-text string (or "" on failure).
# ════════════════════════════════════════════════════════════════

def extract_text_from_txt(filepath: Path) -> str:
    try:
        return filepath.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        log.warning(f"[TXT] {filepath.name}: {e}")
        return ""

def extract_text_from_pdf(filepath: Path) -> str:
    if not HAS_PYPDF2:
        return ""
    try:
        parts = []
        with open(filepath, "rb") as fh:
            reader = PyPDF2.PdfReader(fh)
            for page in reader.pages:
                parts.append(page.extract_text() or "")
        return " ".join(parts)
    except Exception as e:
        log.warning(f"[PDF] {filepath.name}: {e}")
        return ""

def extract_text_from_docx(filepath: Path) -> str:
    if not HAS_DOCX:
        return ""
    try:
        doc = DocxDocument(str(filepath))
        return " ".join(p.text for p in doc.paragraphs)
    except Exception as e:
        log.warning(f"[DOCX] {filepath.name}: {e}")
        return ""

def extract_text_from_pptx(filepath: Path) -> str:
    if not HAS_PPTX:
        return ""
    try:
        prs = PptxPresentation(str(filepath))
        parts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    parts.append(shape.text)
        return " ".join(parts)
    except Exception as e:
        log.warning(f"[PPTX] {filepath.name}: {e}")
        return ""

def extract_text_from_xlsx(filepath: Path) -> str:
    if not HAS_OPENPYXL:
        return ""
    try:
        wb = openpyxl.load_workbook(str(filepath), read_only=True, data_only=True)
        ws = wb.active
        parts = [str(cell) for row in ws.iter_rows(values_only=True)
                 for cell in row if cell is not None]
        return " ".join(parts)
    except Exception as e:
        log.warning(f"[XLSX] {filepath.name}: {e}")
        return ""

def extract_text(filepath: Path, max_chars: int = 8000) -> str:
    """
    Dispatch to the correct extractor based on extension.
    Truncates result to max_chars to keep embeddings fast and consistent.
    """
    ext = filepath.suffix.lower()
    dispatch = {
        ".txt":  extract_text_from_txt,
        ".md":   extract_text_from_txt,
        ".rtf":  extract_text_from_txt,
        ".pdf":  extract_text_from_pdf,
        ".docx": extract_text_from_docx,
        ".pptx": extract_text_from_pptx,
        ".xlsx": extract_text_from_xlsx,
        ".csv":  extract_text_from_txt,
    }
    extractor = dispatch.get(ext)
    raw = extractor(filepath) if extractor else ""
    return raw[:max_chars]


# ════════════════════════════════════════════════════════════════
#  SECTION 4 — SEMANTIC ANCHOR BUILDER
#
#  A "semantic anchor" is the rich text representation of each
#  class folder that we encode into a vector.  The richer it is,
#  the better the model can distinguish between classes.
# ════════════════════════════════════════════════════════════════

def expand_folder_name(folder_name: str) -> str:
    """
    Turn a terse folder name into a more descriptive phrase.

    Strategy:
      • Split on common separators (underscore, hyphen, camelCase)
      • Separate a leading course-code prefix from trailing digits
        e.g. "CS101"   → "CS 101 computer science"
             "MATH203" → "MATH 203 mathematics"
             "HIST110" → "HIST 110 history"
             "ENG201"  → "ENG 201 english"
             "BIO305"  → "BIO 305 biology"
      • Common abbreviation map for expansion
    """
    abbrev_map = {
        "CS":    "computer science programming software",
        "MATH":  "mathematics calculus algebra",
        "HIST":  "history civilization society",
        "ENG":   "english literature writing composition",
        "PHYS":  "physics mechanics energy",
        "CHEM":  "chemistry elements reactions",
        "BIO":   "biology cells organisms genetics",
        "ECON":  "economics market finance",
        "PSYC":  "psychology behavior mind",
        "PHIL":  "philosophy logic ethics",
        "SOC":   "sociology society culture",
        "POLS":  "political science government",
        "GEOG":  "geography spatial environment",
        "STAT":  "statistics data probability",
        "ART":   "art design visual creative",
        "MUS":   "music theory composition sound",
        "NSEC":  "network security cybersecurity",
        "NSSE":  "network security forensics",
        "IT":    "information technology systems",
    }

    # Separate letters from trailing digits: "CS101" → ("CS", "101")
    match = re.match(r"^([A-Za-z]+)(\d+)?(.*)$", folder_name)
    if match:
        prefix   = match.group(1).upper()
        number   = match.group(2) or ""
        suffix   = match.group(3) or ""
        expansion = abbrev_map.get(prefix, prefix.lower())
        return f"{prefix} {number} {suffix} {expansion}".strip()

    # Fallback: split on underscores/hyphens and return as-is
    return re.sub(r"[_\-]+", " ", folder_name)


def build_class_anchor(class_dir: Path) -> str:
    """
    Build a rich semantic anchor string for a class folder.

    Combines:
      1. The expanded folder name (e.g. "MATH 203 mathematics calculus")
      2. Short text samples from existing files already in the folder
         (up to MAX_ANCHOR_FILES files, MAX_ANCHOR_CHARS each)

    The resulting string is what we embed to represent this class.
    """
    folder_name = class_dir.name
    parts = [expand_folder_name(folder_name)]

    if MAX_ANCHOR_FILES > 0 and class_dir.exists():
        sampled = 0
        for f in class_dir.iterdir():
            if sampled >= MAX_ANCHOR_FILES:
                break
            if not f.is_file() or f.name.startswith("."):
                continue
            snippet = extract_text(f, max_chars=MAX_ANCHOR_CHARS)
            if snippet.strip():
                parts.append(snippet)
                sampled += 1

    return " ".join(parts)


# ════════════════════════════════════════════════════════════════
#  SECTION 5 — AI CLASSIFIER  (the heart of the new logic)
# ════════════════════════════════════════════════════════════════

class SemanticClassifier:
    """
    Wraps a SentenceTransformer model and exposes a classify() method.

    On construction:
      • Loads the embedding model (cached locally after first download)
      • Scans CLASSES_FOLDER for subfolders  →  dynamic class list
      • Builds and encodes a semantic anchor for every class
      • Optionally initialises KeyBERT for Unsorted label generation

    classify(filepath) returns a dict describing the best action.
    """

    def __init__(self, classes_dir: Path):
        if not HAS_SBERT:
            log.error("sentence-transformers is not installed.")
            log.error("Run:  pip install sentence-transformers")
            sys.exit(1)

        print(f"\n  Loading embedding model: {EMBEDDING_MODEL}")
        print("  (First run downloads ~80 MB — subsequent runs are instant)\n")
        self.model = SentenceTransformer(EMBEDDING_MODEL)

        # Optionally set up KeyBERT (reuses the same model object)
        if HAS_KEYBERT:
            self.kw_model = KeyBERT(model=self.model)
        else:
            self.kw_model = None
            log.warning("KeyBERT not installed — Unsorted folders will use a simple fallback label.")
            log.warning("Install with:  pip install keybert")

        # ── Discover class folders ────────────────────────────
        self.class_dirs = {
            d.name: d for d in classes_dir.iterdir()
            if d.is_dir() and not d.name.startswith(".")
               and d.name not in {"Review", "Unsorted"}
        }

        if not self.class_dirs:
            log.error(f"No class subfolders found in: {classes_dir}")
            log.error("Please create at least one class subfolder first.")
            sys.exit(1)

        print(f"  Discovered {len(self.class_dirs)} class folder(s):")
        for name in sorted(self.class_dirs):
            print(f"    • {name}")
        print()

        # ── Build and encode class anchors ───────────────────
        print("  Building semantic anchors for each class…")
        self.class_names   = sorted(self.class_dirs.keys())
        anchor_texts       = []
        self.anchor_debug  = {}   # store for reporting

        for name in self.class_names:
            anchor = build_class_anchor(self.class_dirs[name])
            anchor_texts.append(anchor)
            self.anchor_debug[name] = anchor[:120] + "…" if len(anchor) > 120 else anchor

        # Encode all class anchors at once (batched = fast)
        self.class_embeddings = self.model.encode(
            anchor_texts,
            convert_to_tensor=True,
            show_progress_bar=False,
        )
        print("  Semantic anchors ready.\n")

    # ── Public interface ──────────────────────────────────────

    def classify(self, filepath: Path) -> dict:
        """
        Classify a single file.

        Returns a dict:
            class_name   : str | None
            similarity   : float
            action       : "move" | "review" | "unsorted"
            confidence   : "STRONG" | "LOW" | "AMBIGUOUS" | "NONE"
            basis        : human-readable explanation
            unsorted_cat : str | None
        """
        ext = filepath.suffix.lower()

        # Extension-based shortcuts (no text to embed)
        if ext in IMAGE_EXTENSIONS:
            return self._unsorted_result(filepath, "", "Images", "Image file — no text to embed")
        if ext in ARCHIVE_EXTENSIONS:
            return self._unsorted_result(filepath, "", "Archives", "Archive file — no text to embed")

        # ── Build file corpus (filename + extracted text) ─────
        filename_hint = filepath.stem.replace("_", " ").replace("-", " ")
        content       = extract_text(filepath)
        corpus        = (filename_hint + " " + content).strip()

        if not corpus:
            return self._unsorted_result(filepath, corpus, "Misc",
                                         "Empty or unreadable file — no text extracted")

        # ── Embed the file corpus ─────────────────────────────
        file_embedding = self.model.encode(corpus, convert_to_tensor=True)

        # ── Compute cosine similarity vs all class anchors ────
        similarities = util.cos_sim(file_embedding, self.class_embeddings)[0]
        # similarities is a 1-D tensor, one score per class

        # Convert to plain list of (class_name, score) tuples
        scores = [
            (self.class_names[i], float(similarities[i]))
            for i in range(len(self.class_names))
        ]
        scores.sort(key=lambda x: x[1], reverse=True)

        best_name,  best_score  = scores[0]
        second_name, second_score = scores[1] if len(scores) > 1 else ("—", 0.0)

        # ── Apply threshold logic ─────────────────────────────

        # 1. Score too low → no meaningful match
        if best_score < MIN_SIMILARITY:
            cat = self._infer_unsorted_category(filepath, corpus)
            return self._unsorted_result(
                filepath, corpus, cat,
                f"Best similarity {best_score:.3f} < threshold {MIN_SIMILARITY} "
                f"(closest class: {best_name})"
            )

        # 2. Top two classes are too close → ambiguous
        if (best_score - second_score) <= AMBIGUITY_GAP and second_score >= MIN_SIMILARITY:
            return {
                "class_name":   None,
                "similarity":   best_score,
                "action":       "review",
                "confidence":   "AMBIGUOUS",
                "basis":        (f"Ambiguous: {best_name} ({best_score:.3f}) vs "
                                 f"{second_name} ({second_score:.3f}) — gap {best_score - second_score:.3f} "
                                 f"≤ AMBIGUITY_GAP {AMBIGUITY_GAP}"),
                "unsorted_cat": None,
            }

        # 3. Clear winner — determine confidence label
        confidence = "STRONG" if best_score >= STRONG_THRESHOLD else "LOW"
        return {
            "class_name":   best_name,
            "similarity":   best_score,
            "action":       "move",
            "confidence":   confidence,
            "basis":        (f"{confidence} match → {best_name} "
                             f"(similarity={best_score:.3f}, "
                             f"runner-up={second_name} at {second_score:.3f})"),
            "unsorted_cat": None,
        }

    # ── Internal helpers ──────────────────────────────────────

    def _infer_unsorted_category(self, filepath: Path, corpus: str) -> str:
        """
        Use KeyBERT to extract the top 1-2 keywords from the file's
        corpus and turn them into a PascalCase folder name.

        Falls back to a sanitised version of the filename stem if
        KeyBERT is unavailable or extraction fails.
        """
        if self.kw_model and corpus.strip():
            try:
                keywords = self.kw_model.extract_keywords(
                    corpus,
                    keyphrase_ngram_range=(1, UNSORTED_MAX_KEYPHRASE_LEN),
                    stop_words="english",
                    top_n=UNSORTED_KEYPHRASE_COUNT,
                    use_mmr=True,   # Maximal Marginal Relevance → diverse keywords
                    diversity=0.5,
                )
                # keywords = [("tax return", 0.78), ("income", 0.61)]
                if keywords:
                    label_parts = []
                    for kw, _ in keywords:
                        # PascalCase each word in the keyphrase
                        label_parts.append("_".join(w.capitalize() for w in kw.split()))
                    return "_".join(label_parts)   # e.g. "Tax_Returns_Income"
            except Exception as e:
                log.warning(f"KeyBERT extraction failed for {filepath.name}: {e}")

        # Fallback: sanitise the filename stem
        stem = re.sub(r"[^a-zA-Z0-9 ]", " ", filepath.stem)
        words = stem.split()[:3]
        return "_".join(w.capitalize() for w in words) if words else "Misc"

    @staticmethod
    def _unsorted_result(filepath, corpus, category, reason) -> dict:
        return {
            "class_name":   None,
            "similarity":   0.0,
            "action":       "unsorted",
            "confidence":   "NONE",
            "basis":        f"{reason} → Unsorted/{category}",
            "unsorted_cat": category,
        }


# ════════════════════════════════════════════════════════════════
#  SECTION 6 — FILE SYSTEM HELPERS
# ════════════════════════════════════════════════════════════════

def safe_resolve(raw_path: str) -> Path:
    return Path(os.path.expanduser(raw_path)).resolve()

def is_hidden_or_system(filepath: Path) -> bool:
    name = filepath.name
    return name.startswith(".") or name in {"Thumbs.db", "desktop.ini", "__MACOSX"}

def unique_destination(dest_path: Path) -> Path:
    """Append _1, _2, … until the path is free — never overwrites."""
    if not dest_path.exists():
        return dest_path
    counter = 1
    stem, suffix, parent = dest_path.stem, dest_path.suffix, dest_path.parent
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1

def ensure_directory(path: Path, dry_run: bool) -> None:
    if not dry_run:
        path.mkdir(parents=True, exist_ok=True)

def move_file(src: Path, dest: Path, dry_run: bool) -> Path:
    if dry_run:
        log.info(f"  [DRY RUN] Would move: {src.name}  →  {dest}")
        return dest
    shutil.move(str(src), str(dest))
    log.info(f"  [MOVED]   {src.name}  →  {dest}")
    return dest


# ════════════════════════════════════════════════════════════════
#  SECTION 7 — MAIN ORCHESTRATOR
# ════════════════════════════════════════════════════════════════

def run_organizer():
    source_dir  = safe_resolve(SOURCE_FOLDER)
    classes_dir = safe_resolve(CLASSES_FOLDER)

    # ── Banner ────────────────────────────────────────────────
    print("\n" + "═" * 62)
    print("  ClassSort AI — Semantic Academic File Organizer")
    print("═" * 62)
    print(f"  Mode       : {'🔵 DRY RUN (no files will move)' if DRY_RUN else '🟢 LIVE RUN'}")
    print(f"  Source     : {source_dir}")
    print(f"  Classes    : {classes_dir}")
    print(f"  Model      : {EMBEDDING_MODEL}")
    print("═" * 62)

    # ── Validate ──────────────────────────────────────────────
    for label, path in [("Source", source_dir), ("Classes", classes_dir)]:
        if not path.exists():
            log.error(f"{label} folder not found: {path}")
            sys.exit(1)

    # ── Initialise AI classifier ──────────────────────────────
    classifier = SemanticClassifier(classes_dir)

    # ── Collect files ─────────────────────────────────────────
    try:
        all_files = [
            f for f in source_dir.iterdir()
            if f.is_file() and not is_hidden_or_system(f)
        ]
    except PermissionError as e:
        log.error(f"Cannot read source folder: {e}")
        sys.exit(1)

    if not all_files:
        print("\n  No files found in the source folder. Nothing to do.\n")
        return

    print(f"\n  Found {len(all_files)} file(s) to process.\n")

    # ── Special destination folders ───────────────────────────
    review_dir   = classes_dir / "Review"
    unsorted_dir = classes_dir / "Unsorted"

    # ── Stats ─────────────────────────────────────────────────
    results = []
    stats   = defaultdict(int)
    created_unsorted_folders = set()

    # ── Process each file ─────────────────────────────────────
    for filepath in sorted(all_files):
        print(f"  ▶  {filepath.name}")
        try:
            result = classifier.classify(filepath)
        except Exception:
            log.error(f"Unexpected error classifying {filepath.name}:\n{traceback.format_exc()}")
            result = {
                "class_name":   None,
                "similarity":   0.0,
                "action":       "unsorted",
                "confidence":   "NONE",
                "basis":        "Classification error — moved to Misc",
                "unsorted_cat": "Misc",
            }

        action       = result["action"]
        class_name   = result["class_name"]
        basis        = result["basis"]
        similarity   = result["similarity"]
        confidence   = result["confidence"]
        unsorted_cat = result.get("unsorted_cat")

        # ── Determine destination directory ───────────────────
        if action == "move" and class_name:
            dest_dir = classes_dir / class_name
        elif action == "review":
            dest_dir = review_dir
        else:
            sub = unsorted_cat or "Misc"
            dest_dir = unsorted_dir / sub
            created_unsorted_folders.add(sub)

        ensure_directory(dest_dir, DRY_RUN)

        # ── Safe move ─────────────────────────────────────────
        raw_dest  = dest_dir / filepath.name
        safe_dest = unique_destination(raw_dest)
        actual_dest = move_file(filepath, safe_dest, DRY_RUN)

        stats[action]    += 1
        stats["scanned"] += 1

        dest_label = class_name or ("Review" if action == "review" else f"Unsorted/{unsorted_cat}")
        print(f"     → {action.upper()} [{confidence}]  sim={similarity:.3f}  dest={dest_label}")
        print(f"       {basis}\n")

        results.append({
            "filename":       filepath.name,
            "original_path":  str(filepath),
            "dest_folder":    str(dest_dir),
            "final_path":     str(actual_dest),
            "action":         action.upper(),
            "confidence":     confidence,
            "similarity":     f"{similarity:.4f}",
            "destination":    dest_label,
            "basis":          basis,
        })

    # ── Report ────────────────────────────────────────────────
    generate_report(results, stats, created_unsorted_folders, classes_dir,
                    classifier.anchor_debug)


# ════════════════════════════════════════════════════════════════
#  SECTION 8 — REPORTING
# ════════════════════════════════════════════════════════════════

def generate_report(results, stats, created_unsorted_folders,
                    classes_dir, anchor_debug):
    timestamp  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mode_label = "DRY RUN" if DRY_RUN else "LIVE RUN"

    # ── Console summary ───────────────────────────────────────
    div = "─" * 62
    print(div)
    print(f"  CLASSORT AI SUMMARY  [{mode_label}]  {timestamp}")
    print(div)
    print(f"  {'Files scanned':<32} {stats['scanned']:>5}")
    print(f"  {'Moved to class folders':<32} {stats['move']:>5}")
    print(f"  {'Sent to Review (ambiguous)':<32} {stats['review']:>5}")
    print(f"  {'Placed in Unsorted':<32} {stats['unsorted']:>5}")

    if created_unsorted_folders:
        print(f"\n  Unsorted subfolders created:")
        for folder in sorted(created_unsorted_folders):
            count = sum(1 for r in results
                        if r["action"] == "UNSORTED"
                        and r["final_path"].find(f"/{folder}/") != -1)
            print(f"    • Unsorted/{folder:<24} ({count} file{'s' if count != 1 else ''})")

    missing = []
    if not HAS_PYPDF2:   missing.append("PyPDF2")
    if not HAS_DOCX:     missing.append("python-docx")
    if not HAS_PPTX:     missing.append("python-pptx")
    if not HAS_OPENPYXL: missing.append("openpyxl")
    if not HAS_KEYBERT:  missing.append("keybert")
    if missing:
        print(f"\n  ⚠  Missing optional packages:")
        print(f"     pip install {' '.join(missing)}")
    print(div + "\n")

    # ── Text report ───────────────────────────────────────────
    lines = ["=" * 72,
             "  CLASSORT AI — DETAILED SORTING REPORT",
             f"  Generated : {timestamp}   Mode: {mode_label}",
             f"  Source    : {SOURCE_FOLDER}",
             f"  Classes   : {CLASSES_FOLDER}",
             f"  Model     : {EMBEDDING_MODEL}",
             "=" * 72,
             f"  Total scanned   : {stats['scanned']}",
             f"  Moved to class  : {stats['move']}",
             f"  Sent to Review  : {stats['review']}",
             f"  Placed Unsorted : {stats['unsorted']}",
             "",
             "─" * 72,
             "  CLASS SEMANTIC ANCHORS (what the model 'sees' for each class)",
             "─" * 72]
    for name, anchor_preview in sorted(anchor_debug.items()):
        lines.append(f"  {name:<16} {anchor_preview}")

    lines += ["", "─" * 72, "  PER-FILE DETAILS", "─" * 72]
    for r in results:
        lines += [
            f"\n  File       : {r['filename']}",
            f"  From       : {r['original_path']}",
            f"  To         : {r['final_path']}",
            f"  Action     : {r['action']}  [{r['confidence']}]",
            f"  Similarity : {r['similarity']}",
            f"  Destination: {r['destination']}",
            f"  Basis      : {r['basis']}",
        ]
    lines.append("\n" + "=" * 72)
    report_text = "\n".join(lines)

    txt_path = classes_dir / REPORT_TXT_FILENAME
    csv_path = classes_dir / REPORT_CSV_FILENAME

    if not DRY_RUN:
        try:
            txt_path.write_text(report_text, encoding="utf-8")
            print(f"  📄 Report saved → {txt_path}")
        except Exception as e:
            log.warning(f"Could not write .txt report: {e}")
        try:
            fields = ["filename", "original_path", "dest_folder", "final_path",
                      "action", "confidence", "similarity", "destination", "basis"]
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fields)
                writer.writeheader()
                writer.writerows(results)
            print(f"  📊 CSV saved     → {csv_path}\n")
        except Exception as e:
            log.warning(f"Could not write .csv report: {e}")
    else:
        print(f"  [DRY RUN] Would write report → {txt_path}")
        print(f"  [DRY RUN] Would write CSV    → {csv_path}\n")


# ════════════════════════════════════════════════════════════════
#  SECTION 9 — ENTRY POINT
# ════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    run_organizer()