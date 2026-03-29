#!/usr/bin/env python3
"""
════════════════════════════════════════════════════════════════════
  ClassSort AI v5 — Ollama + PyQt6 | Clean UI · Fast · Context-Aware
  macOS | Python 3.9+ | PyQt6 + Ollama (local LLM)
════════════════════════════════════════════════════════════════════

WHAT'S NEW IN v5
  ┌─ UI ──────────────────────────────────────────────────────────┐
  │  • Removed fake nav list, user profile row, dummy tab bar,    │
  │    filter/column buttons, and the Sparkline widget entirely.  │
  │  • Table is now 5 clean columns: # · File · Score · Dest · Type │
  └───────────────────────────────────────────────────────────────┘
  ┌─ AI ──────────────────────────────────────────────────────────┐
  │  • Folder profiling: up to 10 existing filenames per class    │
  │    folder are injected into the system prompt so the LLM can  │
  │    pattern-match incoming files against what's already there. │
  └───────────────────────────────────────────────────────────────┘
  ┌─ SPEED ────────────────────────────────────────────────────────┐
  │  • ThreadPoolExecutor pre-extracts text in parallel before    │
  │    the Ollama batching loop starts.                           │
  │  • MAX_EXTRACT_CHARS = 1500  (was 6000).                      │
  │  • Fast-fail: media/archive extensions skip extraction.       │
  │  • Top-level subfolders in source → instant Unsorted (0 %).   │
  │  • BATCH_SIZE = 10 Ollama calls per round-trip.               │
  └───────────────────────────────────────────────────────────────┘

SETUP
  pip install PyQt6 PyPDF2 python-docx python-pptx openpyxl requests
  brew install ollama && ollama pull llama3.2
  ollama serve          ← keep running in a separate terminal tab
"""

# ════════════════════════════════════════════════════════════════════
#  §1  CONFIGURATION  — edit this block freely
# ════════════════════════════════════════════════════════════════════

SOURCE_FOLDER  = "/Users/justinevaldes/Desktop/toSort"
CLASSES_FOLDER = "/Users/justinevaldes/Desktop/school"

# True  = nothing moves; reports still written.  False = live.
DRY_RUN = True

# ── Ollama ────────────────────────────────────────────────────────
OLLAMA_URL     = "http://localhost:11434/api/generate"
OLLAMA_MODEL   = "llama3.2"       # swap to phi3, mistral, etc.
OLLAMA_TIMEOUT = 90               # seconds per batch request

# ── Confidence tiers ──────────────────────────────────────────────
HIGH_CONFIDENCE   = 75   # ≥ this → green / auto-move
MEDIUM_CONFIDENCE = 40   # ≥ this → yellow / review; < 40 → red / Unsorted

# ── Speed settings ────────────────────────────────────────────────
MAX_EXTRACT_CHARS   = 1500   # keep prompts short for local LLMs
BATCH_SIZE          = 10     # files per Ollama call
EXTRACTION_WORKERS  = 6      # ThreadPoolExecutor thread count

# ── Folder profiling (context-aware AI) ───────────────────────────
MAX_PROFILE_FILES   = 10     # existing filenames to sample per class folder

# ── Report ────────────────────────────────────────────────────────
REPORT_CSV = "sorting_report.csv"


# ════════════════════════════════════════════════════════════════════
#  §2  IMPORTS
# ════════════════════════════════════════════════════════════════════

import os, re, sys, csv, json, shutil, logging, traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from datetime import datetime
from collections import defaultdict

import requests

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QProgressBar, QHeaderView, QFrame, QSizePolicy,
    QMessageBox, QAbstractItemView, QAbstractScrollArea,
    QGraphicsDropShadowEffect,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QRectF
from PyQt6.QtGui  import (
    QColor, QPainter, QPainterPath, QBrush, QPen, QLinearGradient, QFont,
)

# ── Optional extraction libraries ─────────────────────────────────
try:
    import PyPDF2;                        HAS_PYPDF2   = True
except ImportError:                       HAS_PYPDF2   = False
try:
    from docx import Document as DocxDoc; HAS_DOCX     = True
except ImportError:                       HAS_DOCX     = False
try:
    from pptx import Presentation as Prs; HAS_PPTX     = True
except ImportError:                       HAS_PPTX     = False
try:
    import openpyxl;                      HAS_OPENPYXL = True
except ImportError:                       HAS_OPENPYXL = False

logging.basicConfig(level=logging.INFO, format="%(levelname)-8s %(message)s")
log = logging.getLogger("ClassSort-v5")


# ════════════════════════════════════════════════════════════════════
#  §3  TEXT EXTRACTION
#      Fast-fail for media/archives; parallel dispatch via executor.
# ════════════════════════════════════════════════════════════════════

# Extensions that have no readable text — skip immediately
SKIP_EXTENSIONS = {
    # Images
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif",
    ".heic", ".heif", ".webp", ".svg", ".ico", ".raw", ".cr2", ".nef",
    # Video
    ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".m4v", ".webm",
    # Audio
    ".mp3", ".wav", ".aac", ".flac", ".ogg", ".m4a", ".wma", ".aiff",
    # Archives & disk images
    ".zip", ".tar", ".gz", ".bz2", ".xz", ".rar", ".7z",
    ".dmg", ".pkg", ".iso", ".img",
    # Compiled / binary
    ".exe", ".dll", ".so", ".dylib", ".bin", ".o", ".class",
    # Fonts
    ".ttf", ".otf", ".woff", ".woff2",
}

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
        for slide in Prs(str(p)).slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"): parts.append(shape.text)
        return " ".join(parts)
    except: return ""

def _read_xlsx(p: Path) -> str:
    if not HAS_OPENPYXL: return ""
    try:
        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        return " ".join(
            str(c) for row in wb.active.iter_rows(values_only=True)
            for c in row if c is not None
        )
    except: return ""

_EXTRACTORS = {
    ".txt": _read_txt, ".md": _read_txt, ".rtf": _read_txt, ".csv": _read_txt,
    ".log": _read_txt, ".json": _read_txt, ".xml": _read_txt, ".yaml": _read_txt,
    ".py": _read_txt,  ".js": _read_txt,  ".ts": _read_txt,  ".html": _read_txt,
    ".pdf": _read_pdf, ".docx": _read_docx, ".pptx": _read_pptx, ".xlsx": _read_xlsx,
}

def extract_text(p: Path) -> str:
    """
    Return up to MAX_EXTRACT_CHARS of readable text from p.
    Returns "" immediately for SKIP_EXTENSIONS — no I/O touched.
    """
    if p.suffix.lower() in SKIP_EXTENSIONS:
        return ""
    fn = _EXTRACTORS.get(p.suffix.lower())
    return (fn(p) if fn else "")[:MAX_EXTRACT_CHARS]


def parallel_extract(files: list[Path]) -> dict[Path, str]:
    """
    Run extract_text() for every file in parallel using a thread pool.
    Returns {filepath: extracted_text} dict.
    Called once at scan start — results are cached before Ollama batching.
    """
    results: dict[Path, str] = {}
    with ThreadPoolExecutor(max_workers=EXTRACTION_WORKERS) as ex:
        future_map = {ex.submit(extract_text, f): f for f in files}
        for future in as_completed(future_map):
            fp = future_map[future]
            try:    results[fp] = future.result()
            except: results[fp] = ""
    return results


# ════════════════════════════════════════════════════════════════════
#  §4  FOLDER PROFILING  — context injected into every Ollama prompt
# ════════════════════════════════════════════════════════════════════

def build_folder_profiles(classes_dir: Path, class_folders: list[str]) -> dict[str, list[str]]:
    """
    For each class folder, collect up to MAX_PROFILE_FILES existing filenames
    (skipping system files).  Returns {folder_name: [filename, ...]}
    Called once at scan start; costs only a single os.scandir per folder.
    """
    profiles: dict[str, list[str]] = {}
    for name in class_folders:
        folder_path = classes_dir / name
        if not folder_path.is_dir():
            profiles[name] = []
            continue
        existing = []
        try:
            for entry in os.scandir(folder_path):
                if entry.is_file() and not entry.name.startswith("."):
                    existing.append(entry.name)
                    if len(existing) >= MAX_PROFILE_FILES:
                        break
        except PermissionError:
            pass
        profiles[name] = existing
    return profiles


def format_folder_list(class_folders: list[str],
                        profiles: dict[str, list[str]]) -> str:
    """
    Build the folder-list section of the Ollama prompt, including
    the existing-file hints so the LLM can pattern-match.

    Example output line:
      CS101 (Existing files: syllabus.pdf, python_hw1.py, midterm.docx)
      MATH203 (No existing files yet)
    """
    lines = []
    for name in class_folders:
        existing = profiles.get(name, [])
        if existing:
            hint = ", ".join(existing)
            lines.append(f"  - {name} (Existing files: {hint})")
        else:
            lines.append(f"  - {name} (No existing files yet)")
    return "\n".join(lines)


# ════════════════════════════════════════════════════════════════════
#  §5  OLLAMA BATCH CLASSIFIER
#      Sends BATCH_SIZE files per request; parses per-file JSON array.
# ════════════════════════════════════════════════════════════════════

# The system prompt now includes the folder-profile hints.
_SYSTEM_PROMPT_TEMPLATE = """\
You are a file classification assistant. Your job is to assign each file
in a batch to the single best-matching destination folder.

Available destination folders (with hints about what is already inside each):
{folder_list}

Rules:
- Use the folder hints (existing filenames) as strong signals.
  If an incoming file looks like something already in a folder, prefer that folder.
- If nothing matches well, use exactly "Unsorted".
- Confidence must be an integer 0-100.
- Reasoning must be ONE concise sentence.

You will receive a JSON array of objects, each with:
  "index"    : integer (preserve it in output)
  "filename" : string
  "content"  : string (extracted text, may be empty)

Respond ONLY with a JSON array — no markdown, no code fences, no extra text.
Each element must have exactly: index, folder, confidence, reasoning.

Example response for a 2-file batch:
[
  {{"index":0,"folder":"CS101","confidence":88,"reasoning":"Python script matching intro CS homework pattern."}},
  {{"index":1,"folder":"Unsorted","confidence":0,"reasoning":"Binary file with no readable content."}}
]
"""

def _safe_folder(folder: str, class_folders: list[str]) -> tuple[str, int, str]:
    """Validate that folder is a known name; fall back to Unsorted."""
    if folder in class_folders or folder == "Unsorted":
        return folder, None, None
    # Model hallucinated — remap
    return "Unsorted", 0, f"Model suggested unknown folder '{folder}' — remapped."


def classify_batch_ollama(
    batch: list[dict],          # [{"index":int, "filepath":Path, "content":str}]
    class_folders: list[str],
    folder_list_str: str,       # pre-formatted folder+hint string
) -> list[dict]:
    """
    Send one batch to Ollama; return list of result dicts aligned to input indices.
    Falls back gracefully on connection errors or bad JSON.
    """
    # Build the user payload (only filename + content go to LLM)
    user_items = [
        {"index": item["index"], "filename": item["filepath"].name,
         "content": item["content"] or "(no extractable text)"}
        for item in batch
    ]
    system_prompt = _SYSTEM_PROMPT_TEMPLATE.format(folder_list=folder_list_str)

    payload = {
        "model":   OLLAMA_MODEL,
        "system":  system_prompt,
        "prompt":  json.dumps(user_items, ensure_ascii=False),
        "stream":  False,
        "options": {"temperature": 0.1, "num_predict": 512},
    }

    # Default: everything unsorted if something goes wrong
    defaults = {
        item["index"]: {
            "folder": "Unsorted", "confidence": 0,
            "reasoning": "Pending — will retry or manual review needed."
        }
        for item in batch
    }

    try:
        resp = requests.post(OLLAMA_URL, json=payload, timeout=OLLAMA_TIMEOUT)
        resp.raise_for_status()
        raw = resp.json().get("response", "").strip()

        # Strip accidental markdown fences
        raw = re.sub(r"^```(?:json)?", "", raw, flags=re.MULTILINE).strip()
        raw = re.sub(r"```$",          "", raw, flags=re.MULTILINE).strip()

        # Find the JSON array — be forgiving about leading/trailing prose
        m = re.search(r"\[.*\]", raw, flags=re.DOTALL)
        if not m:
            raise ValueError(f"No JSON array in response: {raw[:300]}")

        parsed = json.loads(m.group(0))

        for item in parsed:
            idx        = int(item.get("index", -1))
            folder     = str(item.get("folder", "Unsorted")).strip()
            confidence = max(0, min(100, int(item.get("confidence", 0))))
            reasoning  = str(item.get("reasoning", "No reasoning.")).strip()

            # Validate folder name
            if folder not in class_folders and folder != "Unsorted":
                reasoning  = f"Model suggested unknown folder '{folder}'. {reasoning}"
                folder, confidence = "Unsorted", 0

            if idx in defaults:
                defaults[idx] = {"folder": folder, "confidence": confidence,
                                  "reasoning": reasoning}

    except requests.exceptions.ConnectionError:
        for idx in defaults:
            defaults[idx]["reasoning"] = "Ollama not reachable — is 'ollama serve' running?"
    except Exception as e:
        log.warning(f"Ollama batch error: {e}")
        for idx in defaults:
            defaults[idx]["reasoning"] = f"Batch error: {str(e)[:120]}"

    return defaults   # {index: {folder, confidence, reasoning}}


# ════════════════════════════════════════════════════════════════════
#  §6  FILE SYSTEM HELPERS
# ════════════════════════════════════════════════════════════════════

def safe_resolve(p: str) -> Path:
    return Path(os.path.expanduser(p)).resolve()

def is_system(f: Path) -> bool:
    return f.name.startswith(".") or f.name in {"Thumbs.db", "desktop.ini", "__MACOSX"}

def unique_dest(path: Path) -> Path:
    """Append _1, _2, … until a free name is found. Never overwrites."""
    if not path.exists(): return path
    i = 1
    while True:
        c = path.parent / f"{path.stem}_{i}{path.suffix}"
        if not c.exists(): return c
        i += 1

def collect_files(source_dir: Path) -> tuple[list[Path], list[Path]]:
    """
    Returns (top_level_files, subdir_files).
    - top_level_files : files sitting directly in source_dir → sent to Ollama
    - subdir_files    : files inside subdirectories → instant Unsorted (0 %)
    """
    top_level, subdir = [], []
    try:
        for entry in source_dir.iterdir():
            if entry.is_file() and not is_system(entry):
                top_level.append(entry)
            elif entry.is_dir() and not is_system(entry):
                for sub_entry in entry.rglob("*"):
                    if sub_entry.is_file() and not is_system(sub_entry):
                        subdir.append(sub_entry)
    except PermissionError as e:
        log.error(f"Cannot read source folder: {e}")
    return sorted(top_level), sorted(subdir)

def discover_class_folders(classes_dir: Path) -> list[str]:
    reserved = {"Review", "Unsorted"}
    try:
        return sorted(
            d.name for d in classes_dir.iterdir()
            if d.is_dir() and not d.name.startswith(".") and d.name not in reserved
        )
    except PermissionError:
        return []


# ════════════════════════════════════════════════════════════════════
#  §7  BACKGROUND SCAN THREAD
# ════════════════════════════════════════════════════════════════════

class ScanWorker(QThread):
    """
    Runs entirely off the main thread.
    Signals:
      progress(current, total, status_text)
      result_ready(results_list, dest_options_list)
      error(message)
    """
    progress     = pyqtSignal(int, int, str)
    result_ready = pyqtSignal(list, list)
    error        = pyqtSignal(str)

    def __init__(self, source_dir: Path, classes_dir: Path):
        super().__init__()
        self.source_dir  = source_dir
        self.classes_dir = classes_dir

    # ─────────────────────────────────────────────────────────────
    def run(self):
        try:
            # ── 0. Validate paths ─────────────────────────────
            for label, path in [("Source", self.source_dir),
                                 ("Classes", self.classes_dir)]:
                if not path.exists():
                    self.error.emit(f"{label} folder not found:\n{path}")
                    return

            # ── 1. Discover class folders ─────────────────────
            class_folders = discover_class_folders(self.classes_dir)
            if not class_folders:
                self.error.emit(
                    f"No class subfolders found in:\n{self.classes_dir}\n\n"
                    "Create at least one class folder first."
                )
                return
            dest_options = class_folders + ["Unsorted"]

            # ── 2. Build folder profiles (context for AI) ─────
            self.progress.emit(0, 1, "Profiling destination folders…")
            profiles       = build_folder_profiles(self.classes_dir, class_folders)
            folder_list_str = format_folder_list(class_folders, profiles)

            # ── 3. Collect source files ───────────────────────
            self.progress.emit(0, 1, "Collecting source files…")
            top_level_files, subdir_files = collect_files(self.source_dir)
            all_files = top_level_files + subdir_files

            if not all_files:
                self.error.emit("No files found in the source folder.")
                return

            total   = len(all_files)
            results = []
            idx_counter = 0

            # ── 4. Instant-Unsorted for subfolder files ───────
            for filepath in subdir_files:
                try:    rel_path = filepath.relative_to(self.source_dir)
                except: rel_path = Path(filepath.name)

                results.append(self._make_record(
                    idx_counter, filepath, rel_path,
                    folder="Unsorted", confidence=0, tier="unsorted",
                    reasoning="File is inside a subfolder — moved to Unsorted automatically.",
                ))
                idx_counter += 1

            # ── 5. Parallel text extraction (top-level only) ──
            self.progress.emit(0, total, "Extracting text from files (parallel)…")
            extracted = parallel_extract(top_level_files)

            # ── 6. Ollama batching ────────────────────────────
            batch_input = []
            for filepath in top_level_files:
                try:    rel_path = filepath.relative_to(self.source_dir)
                except: rel_path = Path(filepath.name)

                batch_input.append({
                    "index":    idx_counter,   # global index
                    "filepath": filepath,
                    "rel_path": rel_path,
                    "content":  extracted.get(filepath, ""),
                })
                idx_counter += 1

            processed = 0
            for batch_start in range(0, len(batch_input), BATCH_SIZE):
                batch   = batch_input[batch_start: batch_start + BATCH_SIZE]
                names   = ", ".join(b["filepath"].name for b in batch)
                processed += len(batch)
                self.progress.emit(
                    processed, len(top_level_files),
                    f"Batch {batch_start // BATCH_SIZE + 1}: {names[:80]}…"
                )

                ai_results = classify_batch_ollama(batch, class_folders, folder_list_str)

                for item in batch:
                    ai     = ai_results.get(item["index"], {})
                    folder = ai.get("folder", "Unsorted")
                    conf   = ai.get("confidence", 0)
                    reason = ai.get("reasoning", "No reasoning.")

                    if folder == "Unsorted" or conf < MEDIUM_CONFIDENCE:
                        tier, dest = "unsorted", "Unsorted"
                    elif conf < HIGH_CONFIDENCE:
                        tier, dest = "review",   folder
                    else:
                        tier, dest = "auto",     folder

                    results.append(self._make_record(
                        item["index"], item["filepath"], item["rel_path"],
                        folder=dest, confidence=conf, tier=tier,
                        reasoning=reason,
                    ))

            # Sort by original index so UI order matches filesystem order
            results.sort(key=lambda r: r["index"])
            self.result_ready.emit(results, dest_options)

        except Exception:
            self.error.emit(traceback.format_exc())

    # ─────────────────────────────────────────────────────────────
    @staticmethod
    def _make_record(index, filepath, rel_path, folder, confidence, tier, reasoning):
        return {
            "index":        index,
            "filepath":     filepath,
            "rel_path":     rel_path,
            "rel_str":      str(rel_path),
            "filename":     filepath.name,
            "original_path": str(filepath),
            "ai_folder":    folder,
            "current_dest": folder,
            "confidence":   confidence,
            "reasoning":    reasoning,
            "tier":         tier,
            "final_dest":   folder,
            "final_path":   "",
            "action":       "",
        }


# ════════════════════════════════════════════════════════════════════
#  §8  CUSTOM PAINTED WIDGETS
# ════════════════════════════════════════════════════════════════════

class ScorePill(QWidget):
    """
    Rounded pill — green (High ≥75) / amber (Medium 40-74) / red (Low <40).
    Painted with QPainter so border-radius works without QSS hacks.
    """
    _STYLES = {
        "high":   ("#DCFCE7", "#15803D", "High"),
        "medium": ("#FEF3C7", "#B45309", "Medium"),
        "low":    ("#FEE2E2", "#B91C1C", "Low"),
    }

    def __init__(self, confidence: int, parent=None):
        super().__init__(parent)
        if confidence >= HIGH_CONFIDENCE:
            self._bg, self._fg, self._text = self._STYLES["high"]
        elif confidence >= MEDIUM_CONFIDENCE:
            self._bg, self._fg, self._text = self._STYLES["medium"]
        else:
            self._bg, self._fg, self._text = self._STYLES["low"]

        self._conf = confidence
        self.setFixedHeight(26)
        fm = self.fontMetrics()
        # Width for pill text  +  small numeric badge
        badge_text = f"{self._text} · {confidence}%"
        self.setFixedWidth(fm.horizontalAdvance(badge_text) + 28)
        self._badge_text = badge_text

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        path = QPainterPath()
        path.addRoundedRect(QRectF(0, 0, self.width(), self.height()), 7, 7)
        p.fillPath(path, QBrush(QColor(self._bg)))
        p.setPen(QPen(QColor(self._fg)))
        font = QFont()
        font.setPointSize(10)
        font.setWeight(QFont.Weight.Medium)
        p.setFont(font)
        p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self._badge_text)
        p.end()


class StatusDot(QWidget):
    """Green ✓ circle (auto tier) or gray ○ (review/unsorted)."""
    def __init__(self, checked: bool, parent=None):
        super().__init__(parent)
        self.checked = checked
        self.setFixedSize(22, 22)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        cx, cy, r = self.width() / 2, self.height() / 2, 8.5
        if self.checked:
            p.setPen(QPen(QColor("#22C55E"), 1.5))
            p.setBrush(QBrush(QColor("#DCFCE7")))
            p.drawEllipse(QRectF(cx - r, cy - r, 2*r, 2*r))
            p.setPen(QPen(QColor("#16A34A"), 2,
                          Qt.PenStyle.SolidLine,
                          Qt.PenCapStyle.RoundCap,
                          Qt.PenJoinStyle.RoundJoin))
            path = QPainterPath()
            path.moveTo(cx - 3.5, cy)
            path.lineTo(cx - 1,   cy + 3)
            path.lineTo(cx + 4,   cy - 3)
            p.drawPath(path)
        else:
            p.setPen(QPen(QColor("#D1D5DB"), 1.5))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawEllipse(QRectF(cx - r, cy - r, 2*r, 2*r))
        p.end()


class FlatCombo(QComboBox):
    """Minimal flat combobox for the Destination column."""
    def __init__(self, options: list[str], current: str, parent=None):
        super().__init__(parent)
        self.addItems(options)
        idx = self.findText(current)
        if idx >= 0: self.setCurrentIndex(idx)
        self.setFixedHeight(30)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)


# ════════════════════════════════════════════════════════════════════
#  §9  BACKGROUND GRADIENT + WHITE CARD
# ════════════════════════════════════════════════════════════════════

class GradientBackground(QWidget):
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        grad.setColorAt(0.0, QColor("#EDE9FE"))
        grad.setColorAt(0.5, QColor("#F3F0FF"))
        grad.setColorAt(1.0, QColor("#EEF2FF"))
        p.fillRect(self.rect(), QBrush(grad))
        p.end()


class WhiteCard(QWidget):
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), 16, 16)
        p.fillPath(path, QBrush(QColor("#FFFFFF")))
        p.end()


# ════════════════════════════════════════════════════════════════════
#  §10  QSS STYLESHEET
# ════════════════════════════════════════════════════════════════════

QSS = """
* {
    font-family: "SF Pro Text", "Helvetica Neue", sans-serif;
    font-size: 13px;
    color: #111827;
}

/* ── Sidebar ───────────────────────────────────────────── */
#Sidebar {
    background: transparent;
    border-right: 1px solid #F3F4F6;
    min-width: 212px;
    max-width: 212px;
}
#LogoDot {
    color: #7C3AED;
    font-size: 20px;
    font-weight: 900;
}
#LogoLabel {
    color: #111827;
    font-size: 14px;
    font-weight: 700;
    letter-spacing: -0.3px;
}

/* Sidebar stat tiles */
#StatTile {
    background: #F9FAFB;
    border-radius: 9px;
    padding: 1px;
}
#StatNumber { color: #111827; font-size: 20px; font-weight: 700; }
#StatDesc   { color: #9CA3AF; font-size: 10px; }

/* Mode badge */
#ModeBadge {
    border-radius: 6px;
    font-size: 11px;
    font-weight: 600;
    padding: 4px 10px;
    margin: 0 14px;
}

/* Execute button */
QPushButton#ExecBtn {
    background-color: #7C3AED;
    color: #FFFFFF;
    border: none;
    border-radius: 9px;
    font-size: 13px;
    font-weight: 600;
    padding: 10px 0px;
    margin: 4px 14px 18px 14px;
}
QPushButton#ExecBtn:hover    { background-color: #6D28D9; }
QPushButton#ExecBtn:pressed  { background-color: #5B21B6; }
QPushButton#ExecBtn:disabled { background-color: #E5E7EB; color: #9CA3AF; }

/* ── Content area ──────────────────────────────────────── */
#ContentArea { background: transparent; }
#PageTitle   { font-size: 15px; font-weight: 700; letter-spacing: -0.2px; }
#PageSubtitle{ font-size: 11px; color: #9CA3AF; }

/* ── Loading ───────────────────────────────────────────── */
#LoadTitle  { font-size: 16px; font-weight: 700; }
#LoadStatus { font-size: 11px; color: #9CA3AF; }
QProgressBar {
    background: #F3F4F6; border: none;
    border-radius: 4px; height: 6px; font-size: 0px;
}
QProgressBar::chunk { background: #7C3AED; border-radius: 4px; }

/* ── Table ─────────────────────────────────────────────── */
QTableWidget {
    background-color: #FFFFFF;
    border: none;
    border-radius: 12px;
    gridline-color: transparent;
    outline: 0;
    selection-background-color: #F5F3FF;
    selection-color: #111827;
}
QTableWidget::item {
    border: none;
    border-bottom: 1px solid #F9FAFB;
    padding: 0px 6px;
}
QTableWidget::item:selected { background-color: #F5F3FF; }
QHeaderView { background: transparent; }
QHeaderView::section {
    background-color: #FFFFFF;
    color: #9CA3AF;
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 0.4px;
    border: none;
    border-bottom: 1px solid #F3F4F6;
    padding: 9px 6px;
}
QScrollBar:vertical {
    background: transparent; width: 6px; margin: 0;
}
QScrollBar::handle:vertical {
    background: #E5E7EB; border-radius: 3px; min-height: 24px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { height: 6px; }
QScrollBar::handle:horizontal { background: #E5E7EB; border-radius: 3px; }

/* ── FlatCombo ─────────────────────────────────────────── */
FlatCombo, QComboBox {
    background: #F9FAFB;
    border: 1px solid transparent;
    border-radius: 6px;
    padding: 3px 8px;
    color: #374151;
    font-size: 12px;
}
FlatCombo:hover, QComboBox:hover {
    border: 1px solid #DDD6FE;
    background: #F5F3FF;
}
FlatCombo:on, QComboBox:on { border: 1px solid #7C3AED; }
QComboBox::drop-down { border: none; width: 0px; }
QComboBox QAbstractItemView {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    selection-background-color: #F5F3FF;
    selection-color: #5B21B6;
    padding: 3px;
}
"""


# ════════════════════════════════════════════════════════════════════
#  §11  LOADING SCREEN
# ════════════════════════════════════════════════════════════════════

class LoadingScreen(QWidget):
    _SPIN = ["◜", "◝", "◞", "◟"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._idx = 0
        v = QVBoxLayout(self)
        v.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v.setSpacing(14)
        v.setContentsMargins(60, 60, 60, 60)

        self._spin_lbl = QLabel(self._SPIN[0])
        self._spin_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._spin_lbl.setStyleSheet("font-size:32px; color:#7C3AED; background:transparent;")
        v.addWidget(self._spin_lbl)

        title = QLabel("Analysing files with Ollama")
        title.setObjectName("LoadTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("background:transparent;")
        v.addWidget(title)

        self._status = QLabel("Starting…")
        self._status.setObjectName("LoadStatus")
        self._status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._status.setStyleSheet("background:transparent;")
        v.addWidget(self._status)

        v.addSpacing(6)

        self._bar = QProgressBar()
        self._bar.setRange(0, 1)
        self._bar.setValue(0)
        self._bar.setFixedWidth(360)
        self._bar.setFixedHeight(6)
        v.addWidget(self._bar, alignment=Qt.AlignmentFlag.AlignCenter)

        self._counter = QLabel("0 / 0")
        self._counter.setObjectName("LoadStatus")
        self._counter.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._counter.setStyleSheet("background:transparent;")
        v.addWidget(self._counter)

        timer = QTimer(self)
        timer.timeout.connect(self._tick)
        timer.start(110)

    def _tick(self):
        self._idx = (self._idx + 1) % len(self._SPIN)
        self._spin_lbl.setText(self._SPIN[self._idx])

    def update_progress(self, current: int, total: int, text: str):
        self._bar.setRange(0, max(total, 1))
        self._bar.setValue(current)
        self._counter.setText(f"{current} / {total}")
        short = text if len(text) <= 70 else "…" + text[-67:]
        self._status.setText(short)


# ════════════════════════════════════════════════════════════════════
#  §12  SIDEBAR  (cleaned — no fake nav, no user row)
# ════════════════════════════════════════════════════════════════════

class Sidebar(QWidget):
    execute_clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Sidebar")
        self._build()

    def _build(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # ── Logo ──────────────────────────────────────────────
        logo_row = QWidget()
        logo_row.setStyleSheet("background:transparent;")
        lh = QHBoxLayout(logo_row)
        lh.setContentsMargins(16, 22, 16, 18)
        lh.setSpacing(6)

        dot = QLabel("l.")
        dot.setObjectName("LogoDot")
        dot.setStyleSheet("background:transparent;")
        lh.addWidget(dot)

        name = QLabel("ClassSort AI")
        name.setObjectName("LogoLabel")
        name.setStyleSheet("background:transparent;")
        lh.addWidget(name)
        lh.addStretch()
        v.addWidget(logo_row)

        # ── Divider ───────────────────────────────────────────
        div = QFrame()
        div.setFrameShape(QFrame.Shape.HLine)
        div.setStyleSheet("color:#F3F4F6; background:#F3F4F6; max-height:1px;")
        v.addWidget(div)
        v.addSpacing(14)

        # ── Stat tiles ────────────────────────────────────────
        self._stat_refs: dict[str, QLabel] = {}
        for key, label in [("total", "Files found"), ("auto",   "Auto-move"),
                             ("review","Needs review"), ("unsorted","Unsorted")]:
            tile = self._make_tile(label)
            v.addWidget(tile)
            v.addSpacing(4)
            self._stat_refs[key] = tile._val_label

        v.addStretch()

        # ── Divider ───────────────────────────────────────────
        div2 = QFrame()
        div2.setFrameShape(QFrame.Shape.HLine)
        div2.setStyleSheet("color:#F3F4F6; background:#F3F4F6; max-height:1px;")
        v.addWidget(div2)
        v.addSpacing(10)

        # ── Mode badge ────────────────────────────────────────
        mode_badge = QLabel("🔵  DRY RUN" if DRY_RUN else "🟢  LIVE")
        mode_badge.setObjectName("ModeBadge")
        mode_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if DRY_RUN:
            mode_badge.setStyleSheet(
                "background:#EFF6FF; color:#1D4ED8; border-radius:6px; "
                "font-size:11px; font-weight:600; padding:4px 10px; margin:0 14px;"
            )
        else:
            mode_badge.setStyleSheet(
                "background:#F0FDF4; color:#15803D; border-radius:6px; "
                "font-size:11px; font-weight:600; padding:4px 10px; margin:0 14px;"
            )
        v.addWidget(mode_badge)
        v.addSpacing(8)

        # ── Execute button ────────────────────────────────────
        self._exec_btn = QPushButton(
            "⟳  Simulate Moves" if DRY_RUN else "↗  Execute Moves"
        )
        self._exec_btn.setObjectName("ExecBtn")
        self._exec_btn.setEnabled(False)
        self._exec_btn.clicked.connect(self.execute_clicked.emit)
        v.addWidget(self._exec_btn)

    def _make_tile(self, desc: str) -> QWidget:
        tile = QWidget()
        tile.setObjectName("StatTile")
        tile.setStyleSheet(
            "background:#F9FAFB; border-radius:9px; margin:0 14px;"
        )
        h = QHBoxLayout(tile)
        h.setContentsMargins(12, 8, 12, 8)

        desc_lbl = QLabel(desc)
        desc_lbl.setStyleSheet("color:#9CA3AF; font-size:11px; background:transparent;")
        h.addWidget(desc_lbl)
        h.addStretch()

        val_lbl = QLabel("—")
        val_lbl.setStyleSheet(
            "color:#111827; font-size:14px; font-weight:700; background:transparent;"
        )
        h.addWidget(val_lbl)

        tile._val_label = val_lbl   # store ref for updates
        return tile

    def set_stats(self, total: int, auto: int, review: int, unsorted: int):
        self._stat_refs["total"].setText(str(total))
        self._stat_refs["auto"].setText(str(auto))
        self._stat_refs["review"].setText(str(review))
        self._stat_refs["unsorted"].setText(str(unsorted))

    def enable_execute(self, enabled: bool = True):
        self._exec_btn.setEnabled(enabled)


# ════════════════════════════════════════════════════════════════════
#  §13  CONTENT AREA  (cleaned — no fake tabs, no filter bar)
# ════════════════════════════════════════════════════════════════════

class ContentArea(QWidget):
    dest_changed = pyqtSignal(int, str)   # (file_index, new_destination)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ContentArea")
        self._v = QVBoxLayout(self)
        self._v.setContentsMargins(28, 22, 28, 20)
        self._v.setSpacing(0)
        self._combos: dict[int, FlatCombo] = {}
        self._build_header()

    def _build_header(self):
        row = QHBoxLayout()
        row.setSpacing(8)

        icon = QLabel("⊟")
        icon.setStyleSheet("font-size:14px; color:#9CA3AF; background:transparent;")
        row.addWidget(icon)

        title = QLabel("Files")
        title.setObjectName("PageTitle")
        title.setStyleSheet("background:transparent;")
        row.addWidget(title)
        row.addStretch()

        self._subtitle = QLabel("")
        self._subtitle.setObjectName("PageSubtitle")
        self._subtitle.setStyleSheet("background:transparent;")
        row.addWidget(self._subtitle)

        self._v.addLayout(row)
        self._v.addSpacing(20)

    def show_loading(self):
        self._loading = LoadingScreen()
        self._v.addWidget(self._loading, stretch=1)

    def update_loading(self, current: int, total: int, text: str):
        if hasattr(self, "_loading"):
            self._loading.update_progress(current, total, text)

    def show_table(self, files_data: list, dest_options: list):
        if hasattr(self, "_loading"):
            self._loading.setParent(None)
            self._loading.deleteLater()

        self._combos.clear()
        n = len(files_data)
        self._subtitle.setText(f"{n} file{'s' if n != 1 else ''} · click Destination to reassign")

        # ── 5-column table ────────────────────────────────────
        # Columns: # · File · Score · Destination · Type
        cols = ["#", "File", "Score", "Destination", "Type"]
        self._table = QTableWidget(n, len(cols))
        self._table.setHorizontalHeaderLabels(cols)
        self._table.setShowGrid(False)
        self._table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setAlternatingRowColors(False)
        self._table.verticalHeader().setVisible(False)
        self._table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self._table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self._table.setSizeAdjustPolicy(
            QAbstractScrollArea.SizeAdjustPolicy.AdjustToContents
        )
        self._table.verticalHeader().setDefaultSectionSize(52)

        hdr = self._table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)        # #
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)      # File
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)        # Score
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)        # Destination
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)        # Type
        self._table.setColumnWidth(0, 40)
        self._table.setColumnWidth(2, 148)
        self._table.setColumnWidth(3, 200)
        self._table.setColumnWidth(4, 64)

        for row, data in enumerate(files_data):
            self._fill_row(row, data, dest_options)

        self._v.addWidget(self._table, stretch=1)

    def _fill_row(self, row: int, data: dict, dest_options: list):
        # ── Col 0 — row number ────────────────────────────────
        num = QTableWidgetItem(str(row + 1))
        num.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        num.setForeground(QColor("#D1D5DB"))
        self._table.setItem(row, 0, num)

        # ── Col 1 — filename (bold) + reasoning tooltip ───────
        fname_widget = QWidget()
        fname_widget.setStyleSheet("background:transparent;")
        fh = QHBoxLayout(fname_widget)
        fh.setContentsMargins(6, 0, 6, 0)
        fh.setSpacing(4)

        fn_lbl = QLabel(f"<b>{data['filename']}</b>")
        fn_lbl.setStyleSheet("color:#111827; font-size:13px; background:transparent;")
        fn_lbl.setToolTip(
            f"Path: {data['original_path']}\n\nAI Reasoning: {data['reasoning']}"
        )
        fh.addWidget(fn_lbl)

        # Sub-path hint if file is from a subfolder
        if len(data["rel_path"].parts) > 1:
            sub_lbl = QLabel(f"  {data['rel_path'].parent}")
            sub_lbl.setStyleSheet(
                "color:#9CA3AF; font-size:10px; background:transparent;"
            )
            fh.addWidget(sub_lbl)

        fh.addStretch()
        self._table.setCellWidget(row, 1, fname_widget)

        # ── Col 2 — score pill ────────────────────────────────
        pill_wrap = QWidget()
        pill_wrap.setStyleSheet("background:transparent;")
        ph = QHBoxLayout(pill_wrap)
        ph.setContentsMargins(8, 0, 8, 0)
        ph.addWidget(ScorePill(data["confidence"]))
        ph.addStretch()
        self._table.setCellWidget(row, 2, pill_wrap)

        # ── Col 3 — destination combobox ──────────────────────
        combo = FlatCombo(dest_options, data["current_dest"])
        file_idx = data["index"]
        combo.currentTextChanged.connect(
            lambda t, i=file_idx: self.dest_changed.emit(i, t)
        )
        self._table.setCellWidget(row, 3, combo)
        self._combos[file_idx] = combo

        # ── Col 4 — file extension badge ─────────────────────
        ext      = data["filepath"].suffix.upper().lstrip(".") or "—"
        ext_lbl  = QLabel(ext)
        ext_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        ext_lbl.setStyleSheet(
            "color:#6B7280; font-size:10px; font-weight:600; background:transparent;"
        )
        ext_wrap = QWidget()
        ext_wrap.setStyleSheet("background:transparent;")
        ew = QHBoxLayout(ext_wrap)
        ew.setContentsMargins(4, 0, 4, 0)
        ew.addWidget(ext_lbl, alignment=Qt.AlignmentFlag.AlignCenter)
        self._table.setCellWidget(row, 4, ext_wrap)


# ════════════════════════════════════════════════════════════════════
#  §14  MAIN WINDOW
# ════════════════════════════════════════════════════════════════════

class ClassSortWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.source_dir   = safe_resolve(SOURCE_FOLDER)
        self.classes_dir  = safe_resolve(CLASSES_FOLDER)
        self.files_data:  list = []
        self.dest_options: list = []
        self._worker = None

        self.setWindowTitle("ClassSort AI")
        self.resize(1200, 740)
        self.setMinimumSize(940, 580)

        self._build_ui()
        self._start_scan()

    # ── Shell ─────────────────────────────────────────────────
    def _build_ui(self):
        bg = GradientBackground()
        self.setCentralWidget(bg)

        outer = QHBoxLayout(bg)
        outer.setContentsMargins(20, 20, 20, 20)

        card = WhiteCard()
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(48)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(109, 40, 217, 32))
        card.setGraphicsEffect(shadow)

        card_h = QHBoxLayout(card)
        card_h.setContentsMargins(0, 0, 0, 0)
        card_h.setSpacing(0)

        self._sidebar = Sidebar()
        self._sidebar.execute_clicked.connect(self._execute_moves)
        card_h.addWidget(self._sidebar)

        self._content = ContentArea()
        self._content.dest_changed.connect(self._on_dest_changed)
        card_h.addWidget(self._content, stretch=1)

        outer.addWidget(card)

    # ── Scan ──────────────────────────────────────────────────
    def _start_scan(self):
        self._content.show_loading()
        self._worker = ScanWorker(self.source_dir, self.classes_dir)
        self._worker.progress.connect(
            lambda c, t, txt: self._content.update_loading(c, t, txt)
        )
        self._worker.result_ready.connect(self._on_scan_done)
        self._worker.error.connect(self._on_scan_error)
        self._worker.start()

    def _on_scan_error(self, msg: str):
        QMessageBox.critical(self, "ClassSort Error", msg)
        self.close()

    def _on_scan_done(self, results: list, dest_options: list):
        self.files_data   = results
        self.dest_options = dest_options

        auto    = sum(1 for d in results if d["tier"] == "auto")
        review  = sum(1 for d in results if d["tier"] == "review")
        unsorted = sum(1 for d in results if d["tier"] == "unsorted")
        self._sidebar.set_stats(len(results), auto, review, unsorted)
        self._content.show_table(results, dest_options)
        self._sidebar.enable_execute(True)

    def _on_dest_changed(self, index: int, new_dest: str):
        self.files_data[index]["current_dest"] = new_dest

    # ── Execute ───────────────────────────────────────────────
    def _execute_moves(self):
        n    = len(self.files_data)
        mode = "DRY RUN" if DRY_RUN else "LIVE"
        note = ("No files will be moved — simulation only.\n"
                if DRY_RUN else "Files WILL be permanently moved.\n")

        if QMessageBox.question(
            self, "Confirm",
            f"Mode: {mode}\n\n{note}Processing {n} file(s).\nProceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return

        self._sidebar.enable_execute(False)

        unsorted_dir = self.classes_dir / "Unsorted"
        report_rows: list = []
        stats = defaultdict(int)

        for data in self.files_data:
            filepath  = data["filepath"]
            dest_name = data["current_dest"]
            rel_path  = data["rel_path"]

            # Preserve subfolder structure in destination
            sub = rel_path.parent if len(rel_path.parts) > 1 else Path(".")
            dest_dir = (
                unsorted_dir / sub if dest_name == "Unsorted"
                else self.classes_dir / dest_name / sub
            )

            raw_path  = dest_dir / filepath.name
            safe_path = unique_dest(raw_path)

            if DRY_RUN:
                action, final = "SIMULATED", str(safe_path)
            else:
                try:
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(filepath), str(safe_path))
                    action, final = "MOVED", str(safe_path)
                except Exception as e:
                    action, final = "ERROR", f"(failed) {e}"
                    log.error(f"{filepath.name}: {e}")

            data.update({"final_dest": dest_name, "final_path": final, "action": action})
            stats[action] += 1
            report_rows.append({
                "filename":     data["filename"],
                "relative_path": data["rel_str"],
                "original_path": data["original_path"],
                "final_dest":   dest_name,
                "final_path":   final,
                "action":       action,
                "confidence":   data["confidence"],
                "tier":         data["tier"],
                "ai_folder":    data["ai_folder"],
                "reasoning":    data["reasoning"],
            })

        csv_path = self._write_csv(report_rows)
        self._print_summary(report_rows, stats, csv_path)

        QMessageBox.information(
            self, "Complete",
            f"{'Simulation' if DRY_RUN else 'Execution'} complete!\n\n"
            + "\n".join(f"{k}: {v}" for k, v in sorted(stats.items()))
            + f"\n\nReport:\n{csv_path}",
        )
        self.close()

    # ── Report helpers ─────────────────────────────────────────
    def _write_csv(self, rows: list) -> Path:
        path   = self.classes_dir / REPORT_CSV
        fields = ["filename", "relative_path", "original_path", "final_dest",
                  "final_path", "action", "confidence", "tier", "ai_folder", "reasoning"]
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=fields)
                w.writeheader()
                w.writerows(rows)
            log.info(f"CSV → {path}")
        except Exception as e:
            log.warning(f"CSV write failed: {e}")
        return path

    def _print_summary(self, rows: list, stats: dict, csv_path: Path):
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mode = "DRY RUN" if DRY_RUN else "LIVE"
        sep  = "═" * 72

        print(f"\n{sep}")
        print(f"  CLASSORT AI v5 — REPORT  [{mode}]  {ts}")
        print(sep)
        print(f"  Model  : {OLLAMA_MODEL}  (batch={BATCH_SIZE}, "
              f"extract_chars={MAX_EXTRACT_CHARS}, workers={EXTRACTION_WORKERS})")
        print(f"  Source : {self.source_dir}")
        print(f"  Total  : {len(rows)}")
        for k, v in sorted(stats.items()):
            print(f"  {k:<14}: {v}")
        print(f"\n  CSV : {csv_path}")
        print(sep)
        print(f"  {'FILE':<42} {'DEST':<16} {'CONF':>4}  ACTION")
        print("  " + "─" * 68)
        for r in rows:
            fn   = r["filename"][:41]
            dest = r["final_dest"][:15]
            print(f"  {fn:<42} {dest:<16} {r['confidence']:>3}%  {r['action']}")
        print(sep + "\n")


# ════════════════════════════════════════════════════════════════════
#  §15  ENTRY POINT
# ════════════════════════════════════════════════════════════════════

def main():
    # Dependency check
    for pkg in ["PyQt6", "requests"]:
        try:
            __import__(pkg)
        except ImportError:
            print(f"\n[ERROR] Missing: {pkg}  →  pip install {pkg}")
            sys.exit(1)

    # Warn if Ollama isn't reachable (non-fatal — scan will still run)
    try:
        requests.get("http://localhost:11434", timeout=3)
    except Exception:
        print("\n⚠  Ollama not reachable at localhost:11434")
        print("   Start with:  ollama serve")
        print("   Files will receive confidence=0 / Unsorted without it.\n")

    app = QApplication(sys.argv)
    app.setApplicationName("ClassSort AI")
    app.setStyleSheet(QSS)

    window = ClassSortWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()