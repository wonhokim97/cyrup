import sys
import os
import re
import shutil
import tempfile
import traceback
import zipfile
import subprocess
import logging
logging.getLogger("fontTools").setLevel(logging.ERROR)
from dataclasses import dataclass, field, replace
from typing import Optional, Tuple, List, Dict, Any
from datetime import datetime
from pathlib import Path
from html import escape

from PySide6.QtCore import Qt, QSettings, QObject, Signal, Slot, QThread
from PySide6.QtGui import (
    QAction, QActionGroup, QIcon, QKeySequence, QFontDatabase, QFont,
    QDragEnterEvent, QDropEvent
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QLabel, QComboBox, QFileDialog, QMessageBox,
    QPlainTextEdit, QSplitter, QTextBrowser, QStatusBar,
    QSpinBox, QDoubleSpinBox, QCheckBox, QToolButton, QFrame,
    QStyle, QDialog, QDialogButtonBox, QFormLayout, QTabWidget,
    QFontComboBox, QGroupBox, QScrollArea
)

# -------- 라이브러리 체크 --------
try:
    import trafilatura
except Exception:
    trafilatura = None

try:
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos
    if not hasattr(FPDF, "add_font"):
        raise ImportError
except Exception:
    raise RuntimeError(
        "PDF 저장을 위해 fpdf2가 필요합니다.\n"
        "pip install --upgrade fpdf2"
    )

try:
    from ebooklib import epub
except Exception:
    epub = None

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception:
    Image = None

# fontTools(옵션) - 내부 서브셋
try:
    from fontTools.subset import Subsetter, Options as FtOptions  # type: ignore
    from fontTools.ttLib import TTFont  # type: ignore
except Exception:
    Subsetter = None
    FtOptions = None
    TTFont = None


# =========================
# 책 판형(재단) 프리셋 (mm)
# =========================
BOOK_TRIM_SIZES_MM: Dict[str, Optional[Tuple[int, int]]] = {
    "문고판(A6) 105×148": (105, 148),
    "46판(B6) 128×188": (128, 188),
    "변형(130×200) 소설 트렌드": (130, 200),
    "국판 변형(140×205)": (140, 205),
    "국판(A5) 148×210": (148, 210),
    "신국판 152×225": (152, 225),
    "크라운판 176×248": (176, 248),
    "B5(46배판) 188×257": (188, 257),
    "A4 210×297": (210, 297),
    "사용자 지정(mm)": None,
}

# =========================
# 리더 프리셋
# =========================
READER_PRESETS: Dict[str, dict] = {
    "기본(일반)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=82, repack_zip=True, zip_level=9),
    "Kindle(추천)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=80, repack_zip=True, zip_level=9),
    "리디북스(추천)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=82, repack_zip=True, zip_level=9),
    "교보 eBook(추천)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=82, repack_zip=True, zip_level=9),
    "Apple Books(추천)": dict(embed_font=False, cover_max_px=1400, cover_format="JPEG", cover_jpeg_quality=85, repack_zip=True, zip_level=9),
    "폰트임베드(용량↑/호환성↑)": dict(embed_font=True, cover_max_px=1400, cover_format="JPEG", cover_jpeg_quality=85, repack_zip=True, zip_level=9),
}

# ==========================================================
# Async QFileDialog Helper (open() 기반)
# ==========================================================
class AsyncFileDialogs(QObject):
    """
    QFileDialog를 open() + finished 시그널로 사용하기 위한 헬퍼.
    - exec() 대신 open() 사용
    - DontUseNativeDialog 사용(라벨/폰트 적용 안정화)
    - last_dir_key로 마지막 폴더 기억
    """
    def __init__(self, parent: QWidget, settings: QSettings):
        super().__init__(parent)
        self._parent = parent
        self._settings = settings
        self._alive: List[QFileDialog] = []

    def _start_dir(self, key: str) -> str:
        d = self._settings.value(key, "", type=str)
        return d if d else str(Path.home())

    def _remember_dir(self, key: str, file_path: str):
        try:
            self._settings.setValue(key, str(Path(file_path).parent))
        except Exception:
            pass

    def open_existing_file(self, caption: str, name_filter: str, last_dir_key: str, on_done, directory: Optional[str] = None):
        dlg = QFileDialog(self._parent, caption)
        dlg.setFileMode(QFileDialog.ExistingFile)
        dlg.setNameFilter(name_filter)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setDirectory(directory or self._start_dir(last_dir_key))
        self._alive.append(dlg)

        def _finished(result: int):
            try:
                if result == QDialog.Accepted:
                    files = dlg.selectedFiles()
                    chosen = files[0] if files else ""
                    if chosen:
                        self._remember_dir(last_dir_key, chosen)
                    on_done(chosen)
                else:
                    on_done("")
            finally:
                if dlg in self._alive:
                    self._alive.remove(dlg)
                dlg.deleteLater()

        dlg.finished.connect(_finished)
        dlg.open()

    def open_save_file(self, caption: str, default_name: str, name_filter: str, default_suffix: str,
                       last_dir_key: str, on_done, labels_ko: bool = True, directory: Optional[str] = None):
        dlg = QFileDialog(self._parent, caption)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilter(name_filter)
        dlg.setDefaultSuffix(default_suffix)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setDirectory(directory or self._start_dir(last_dir_key))
        dlg.selectFile(default_name)

        if labels_ko:
            dlg.setLabelText(QFileDialog.LookIn, "저장 위치:")
            dlg.setLabelText(QFileDialog.FileName, "파일 이름:")
            dlg.setLabelText(QFileDialog.FileType, "파일 형식:")
            dlg.setLabelText(QFileDialog.Accept, "저장")
            dlg.setLabelText(QFileDialog.Reject, "취소")

        self._alive.append(dlg)

        def _finished(result: int):
            try:
                if result == QDialog.Accepted:
                    files = dlg.selectedFiles()
                    chosen = files[0] if files else ""
                    if chosen:
                        self._remember_dir(last_dir_key, chosen)
                    on_done(chosen)
                else:
                    on_done("")
            finally:
                if dlg in self._alive:
                    self._alive.remove(dlg)
                dlg.deleteLater()

        dlg.finished.connect(_finished)
        dlg.open()


# ==========================================================
# Drag & Drop Text Editor (여러 파일 + 덮어쓰기/이어붙이기)
# ==========================================================
class DropTextEdit(QPlainTextEdit):
    files_dropped = Signal(list)  # List[str]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setPlaceholderText(
            "본문을 입력하거나\n"
            "텍스트 파일(.txt, .md)을 여기로 드래그하세요.\n"
            "여러 파일을 드롭하면 자동 합치기가 가능합니다."
        )

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for u in event.mimeData().urls():
                if u.isLocalFile():
                    p = u.toLocalFile().lower()
                    if p.endswith((".txt", ".md")):
                        event.acceptProposedAction()
                        return
        event.ignore()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        paths = []
        for u in event.mimeData().urls():
            if u.isLocalFile():
                p = u.toLocalFile()
                if p.lower().endswith((".txt", ".md")):
                    paths.append(p)
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()


# =========================
# 텍스트 파서(챕터/문단)
# =========================
CHAPTER_PATTERNS = [
    r'^\s*#{1,6}\s+(.+?)\s*$',
    r'^\s*(제\s*\d+\s*(화|장)|\d+\s*(화|장))\s*[:.\-]?\s*(.*)\s*$',
    r'^\s*(Chapter|CHAPTER|Episode|EPISODE)\s*\d+\s*[:.\-]?\s*(.*)\s*$',
    r'^\s*(프롤로그|에필로그|외전|후기|서장|서문|머리말|끝)\s*[:.\-]?\s*(.*)\s*$',
]
DIALOGUE_PREFIXES = ('"', "'", '“', '‘', '「', '『', '—', '―', '-', '–', '…')
SCENE_BREAK_RE = re.compile(r'^[\s\*\-_=~·•\.]{3,}$')
LISTLIKE_RE = re.compile(r'^\s*(?:\d+[\.\)]|[※▶■◆•\-])\s+')
SENTENCE_END_RE = re.compile(r'.*[\.!?…]["”’\')\]]*\s*$')
KOR_END_RE = re.compile(r'.*(다\.|요\.|죠\.|네\.|임\.|함\.|다!|다\?|요!|요\?)\s*$')


def detect_chapter_title(line: str, max_len: int = 36) -> Optional[str]:
    for pat in CHAPTER_PATTERNS:
        m = re.match(pat, line)
        if not m:
            continue
        groups = [g for g in m.groups() if g is not None and g != ""]
        if pat.startswith(r'^\s*#{'):
            t = groups[0].strip()
            return t if len(t) <= max_len else None

        main = (groups[0].strip() if groups else line.strip())
        sub = (groups[-1].strip() if len(groups) >= 2 else "")
        title = (main + (f" - {sub}" if sub and sub != main else "")).strip()
        return title if len(title) <= max_len else None
    return None


def _count_nonempty(lines: List[str]) -> int:
    return sum(1 for s in lines if s.strip())


def split_into_chapters(text: str, strict: bool = True, min_body_lines: int = 3) -> List[Dict[str, Any]]:
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    chapters: List[Dict[str, Any]] = []
    cur_title = "본문"
    cur_lines: List[str] = []

    for line in lines:
        t = detect_chapter_title(line)
        if t:
            if strict and chapters and _count_nonempty(cur_lines) < min_body_lines:
                cur_lines.append(line)
                continue
            if cur_lines and any(s.strip() for s in cur_lines):
                chapters.append({"title": cur_title, "lines": cur_lines})
            cur_title = t
            cur_lines = []
        else:
            cur_lines.append(line)

    if cur_lines and any(s.strip() for s in cur_lines):
        chapters.append({"title": cur_title, "lines": cur_lines})

    return chapters if chapters else [{"title": "본문", "lines": lines}]


def is_scene_break(line: str) -> bool:
    return bool(SCENE_BREAK_RE.match(line.strip()))


def is_dialogue(line: str) -> bool:
    s = line.lstrip()
    return s.startswith(DIALOGUE_PREFIXES)


def looks_like_list(line: str) -> bool:
    return bool(LISTLIKE_RE.match(line))


def smart_make_paragraphs(lines: List[str]) -> List[str]:
    out: List[str] = []
    buf: List[str] = []

    def flush():
        nonlocal buf
        if buf:
            out.append(" ".join(buf).strip())
            buf = []

    for ln in lines:
        s = ln.rstrip("\n")
        t = s.strip()

        if not t:
            flush()
            out.append("")
            continue

        if SCENE_BREAK_RE.match(t):
            flush()
            out.append("")
            continue

        if LISTLIKE_RE.match(t) or t.lstrip().startswith(DIALOGUE_PREFIXES):
            flush()
            out.append(t)
            continue

        buf.append(t)

    flush()
    # tail empty lines trim
    while out and out[-1] == "":
        out.pop()
    return out

def contains_korean(text: str) -> bool:
    for ch in text:
        o = ord(ch)
        if (0xAC00 <= o <= 0xD7A3) or (0x1100 <= o <= 0x11FF) or (0x3130 <= o <= 0x318F):
            return True
    return False


# =========================
# 프로파일/옵션
# =========================
@dataclass
class PdfFontProfile:
    family: str = "맑은 고딕"
    file_path: str = ""
    body_size_pt: int = 11
    chapter_title_pt: int = 15
    cover_title_pt: int = 24


@dataclass
class EpubFontProfile:
    family: str = "맑은 고딕"
    file_path: str = ""
    base_font_em: float = 1.0
    embed_font: bool = False


@dataclass
class EpubOptimizeOptions:
    optimize_mode: bool = True
    reader_preset: str = "기본(일반)"

    # ✅ 추가: 판형별 추천 타이포 자동 튜닝
    auto_typo_tune: bool = True

    subset_font: bool = True
    prefer_internal_subsetter: bool = True

    cover_max_px: int = 1200
    cover_format: str = "JPEG"
    cover_jpeg_quality: int = 82

    repack_zip: bool = True
    zip_compresslevel: int = 9

    external_epub_tool_enabled: bool = False
    external_epub_tool_cmd: str = ""

    external_font_subset_enabled: bool = False
    external_font_subset_cmd: str = ""


@dataclass
class TypesetOptions:
    # ✅ 판형 확장
    paper: str = "국판(A5) 148×210"
    custom_paper_w_mm: int = 148
    custom_paper_h_mm: int = 210

    line_spacing: float = 1.8
    para_spacing: int = 2
    indent_spaces: int = 4
    show_header: bool = True
    include_title_page: bool = True
    chapter_strict: bool = True
    chapter_min_lines: int = 3
    preview_profile: str = "pdf"

    pdf_profile: PdfFontProfile = field(default_factory=PdfFontProfile)
    epub_profile: EpubFontProfile = field(default_factory=EpubFontProfile)
    epub_opt: EpubOptimizeOptions = field(default_factory=EpubOptimizeOptions)


# =========================
# PDF 클래스
# =========================
class NovelPDF(FPDF):
    def __init__(self, book_title="", show_header=True, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.book_title = book_title
        self.show_header = show_header
        self._body_family = "Helvetica"

    def set_body_font(self, family: str):
        self._body_family = family

    def header(self):
        if not self.show_header or not self.book_title:
            return
        self.set_font(self._body_family, size=8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 8, text=self.book_title, align="C")
        self.ln(10)

    def footer(self):
        self.set_y(-12)
        self.set_font(self._body_family, size=8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 10, text=str(self.page_no()), align="C")


# =========================
# 비동기 추출 워커
# =========================
class ExtractWorker(QObject):
    finished = Signal(str)
    error = Signal(str)
    log_signal = Signal(str)

    def __init__(self, url: str):
        super().__init__()
        self.url = url

    @Slot()
    def run(self):
        if not trafilatura:
            self.error.emit("trafilatura 라이브러리가 설치되지 않았습니다. (pip install trafilatura)")
            return
        try:
            self.log_signal.emit(f"데이터 추출 중: {self.url}")
            downloaded = trafilatura.fetch_url(self.url)
            if not downloaded:
                self.error.emit("URL 접근 실패 (네트워크 또는 사이트 차단)")
                return
            result = trafilatura.extract(downloaded, include_comments=False, include_links=False)
            if result:
                self.finished.emit(result)
            else:
                self.error.emit("본문을 추출할 수 없습니다.")
        except Exception as e:
            self.error.emit(str(e))

# =========================
# 조판 설정 다이얼로그
# =========================
class TypesetDialog(QDialog):
    SAMPLE_TEXT = (
        "미리보기: 가나다라마바사 ABC abc 123\n\n"
        "“대사는 들여쓰기 없이 표시됩니다.”\n"
        "일반 문단은 들여쓰기가 적용됩니다."
    )

    def __init__(self, parent, opts: TypesetOptions, settings: QSettings):
        super().__init__(parent)
        self.setWindowTitle("조판 설정 (통합 + 프리셋 + 최적화)")
        self.resize(920, 720)
        self.setMinimumSize(720, 520)
        self.setSizeGripEnabled(True)
        self.setWindowFlag(Qt.WindowType.Window)

        self.opts = opts
        self.settings = settings
        self.dialogs = AsyncFileDialogs(self, self.settings)

        root = QVBoxLayout(self)
        tabs = QTabWidget()
        root.addWidget(tabs)

        # ==========================================================
        # 1) 기본 조판 탭 (판형 확장 + 커스텀 mm)
        # ==========================================================
        page_basic = QWidget()
        form = QFormLayout(page_basic)

        self.paper_combo = QComboBox()
        self.paper_combo.addItems(list(BOOK_TRIM_SIZES_MM.keys()))

        cur_paper = opts.paper or "국판(A5) 148×210"
        # 하위호환
        if cur_paper == "A5":
            cur_paper = "국판(A5) 148×210"
        elif cur_paper == "A4":
            cur_paper = "A4 210×297"
        if cur_paper in BOOK_TRIM_SIZES_MM:
            self.paper_combo.setCurrentText(cur_paper)
        else:
            self.paper_combo.setCurrentText("국판(A5) 148×210")

        self.custom_w = QSpinBox()
        self.custom_w.setRange(80, 400)
        self.custom_w.setValue(int(getattr(opts, "custom_paper_w_mm", 148)))

        self.custom_h = QSpinBox()
        self.custom_h.setRange(120, 600)
        self.custom_h.setValue(int(getattr(opts, "custom_paper_h_mm", 210)))

        row_custom = QHBoxLayout()
        row_custom.addWidget(QLabel("가로(mm)"))
        row_custom.addWidget(self.custom_w)
        row_custom.addSpacing(10)
        row_custom.addWidget(QLabel("세로(mm)"))
        row_custom.addWidget(self.custom_h)

        self.line_spacing = QDoubleSpinBox()
        self.line_spacing.setRange(1.2, 2.6)
        self.line_spacing.setSingleStep(0.1)
        self.line_spacing.setValue(opts.line_spacing)

        self.para_spacing = QSpinBox()
        self.para_spacing.setRange(0, 24)
        self.para_spacing.setValue(opts.para_spacing)

        self.indent_spaces = QSpinBox()
        self.indent_spaces.setRange(0, 12)
        self.indent_spaces.setValue(opts.indent_spaces)

        self.header_check = QCheckBox("PDF 헤더 표시")
        self.header_check.setChecked(opts.show_header)

        self.cover_check = QCheckBox("표지(Title Page) 포함(PDF/EPUB)")
        self.cover_check.setChecked(opts.include_title_page)

        self.strict_check = QCheckBox("챕터 인식 엄격")
        self.strict_check.setChecked(opts.chapter_strict)

        self.min_lines = QSpinBox()
        self.min_lines.setRange(1, 30)
        self.min_lines.setValue(opts.chapter_min_lines)

        self.preview_profile_combo = QComboBox()
        self.preview_profile_combo.addItems(["PDF 프로파일로 미리보기", "EPUB 프로파일로 미리보기"])
        self.preview_profile_combo.setCurrentIndex(0 if opts.preview_profile == "pdf" else 1)

        form.addRow("용지(판형)", self.paper_combo)
        form.addRow("사용자 지정 크기", row_custom)
        form.addRow("줄간격", self.line_spacing)
        form.addRow("문단간격(px)", self.para_spacing)
        form.addRow("들여쓰기(스페이스)", self.indent_spaces)
        form.addRow("", self.header_check)
        form.addRow("", self.cover_check)
        form.addRow("", self.strict_check)
        form.addRow("최소 본문 줄(엄격모드)", self.min_lines)
        form.addRow("미리보기 기준", self.preview_profile_combo)

        tabs.addTab(page_basic, "기본 조판")

        self.paper_combo.currentTextChanged.connect(self._sync_custom_paper_ui)
        self._sync_custom_paper_ui(self.paper_combo.currentText())

        # ==========================================================
        # 2) PDF 프로파일 탭
        # ==========================================================
        page_pdf = QWidget()
        v_pdf = QVBoxLayout(page_pdf)
        form_pdf = QFormLayout()
        v_pdf.addLayout(form_pdf)

        self.pdf_font_combo = QFontComboBox()
        self.pdf_font_combo.setCurrentFont(QFont(opts.pdf_profile.family))

        self.pdf_font_path = QLineEdit(opts.pdf_profile.file_path)
        btn_pdf_browse = QToolButton(text="파일 지정...")
        btn_pdf_browse.clicked.connect(self._browse_pdf_font)

        row_pdf_path = QHBoxLayout()
        row_pdf_path.addWidget(self.pdf_font_path, 1)
        row_pdf_path.addWidget(btn_pdf_browse)

        self.pdf_body_size = QSpinBox()
        self.pdf_body_size.setRange(8, 24)
        self.pdf_body_size.setValue(opts.pdf_profile.body_size_pt)

        self.pdf_chapter_size = QSpinBox()
        self.pdf_chapter_size.setRange(10, 40)
        self.pdf_chapter_size.setValue(opts.pdf_profile.chapter_title_pt)

        self.pdf_cover_size = QSpinBox()
        self.pdf_cover_size.setRange(12, 60)
        self.pdf_cover_size.setValue(opts.pdf_profile.cover_title_pt)

        form_pdf.addRow("PDF 폰트(패밀리)", self.pdf_font_combo)
        form_pdf.addRow("PDF 폰트 파일(선택)", row_pdf_path)
        form_pdf.addRow("PDF 본문 크기(pt)", self.pdf_body_size)
        form_pdf.addRow("PDF 챕터 제목(pt)", self.pdf_chapter_size)
        form_pdf.addRow("PDF 표지 제목(pt)", self.pdf_cover_size)

        self.pdf_preview = QTextBrowser()
        self.pdf_preview.setPlainText(self.SAMPLE_TEXT)
        self.pdf_preview.setMinimumHeight(150)
        v_pdf.addWidget(QLabel("폰트 미리보기(PDF)"))
        v_pdf.addWidget(self.pdf_preview)
        tabs.addTab(page_pdf, "PDF 프로파일")

        # ==========================================================
        # 3) EPUB 프로파일 탭 + 자동 타이포 튜닝
        # ==========================================================
        page_epub = QWidget()
        v_epub = QVBoxLayout(page_epub)
        form_epub = QFormLayout()
        v_epub.addLayout(form_epub)

        self.epub_font_combo = QFontComboBox()
        self.epub_font_combo.setCurrentFont(QFont(opts.epub_profile.family))

        self.epub_font_path = QLineEdit(opts.epub_profile.file_path)
        btn_epub_browse = QToolButton(text="파일 지정...")
        btn_epub_browse.clicked.connect(self._browse_epub_font)

        row_epub_path = QHBoxLayout()
        row_epub_path.addWidget(self.epub_font_path, 1)
        row_epub_path.addWidget(btn_epub_browse)

        self.epub_base_em = QDoubleSpinBox()
        self.epub_base_em.setRange(0.7, 1.6)
        self.epub_base_em.setSingleStep(0.05)
        self.epub_base_em.setValue(opts.epub_profile.base_font_em)

        self.epub_embed = QCheckBox("EPUB에 폰트 임베드(@font-face)  ※ 용량 증가")
        self.epub_embed.setChecked(bool(opts.epub_profile.embed_font))

        self.epub_auto_tune = QCheckBox("판형별 추천 타이포 프리셋 자동 튜닝(가독성)")
        self.epub_auto_tune.setChecked(bool(opts.epub_opt.auto_typo_tune))

        form_epub.addRow("EPUB 폰트(패밀리)", self.epub_font_combo)
        form_epub.addRow("EPUB 폰트 파일(선택)", row_epub_path)
        form_epub.addRow("EPUB 본문 크기(em)", self.epub_base_em)
        form_epub.addRow("", self.epub_embed)
        form_epub.addRow("", self.epub_auto_tune)

        self.epub_preview = QTextBrowser()
        self.epub_preview.setPlainText(self.SAMPLE_TEXT)
        self.epub_preview.setMinimumHeight(150)
        v_epub.addWidget(QLabel("폰트 미리보기(EPUB)"))
        v_epub.addWidget(self.epub_preview)
        tabs.addTab(page_epub, "EPUB 프로파일")

        # ==========================================================
        # 4) 프리셋 탭 (프리셋 적용 + 설명)
        # ==========================================================
        page_preset = QWidget()
        v_preset = QVBoxLayout(page_preset)
        v_preset.setSpacing(10)

        box_preset = QGroupBox("리더 프리셋")
        form_preset = QFormLayout(box_preset)

        self.preset_combo = QComboBox()
        self.preset_combo.addItems(list(READER_PRESETS.keys()))
        self.preset_combo.setCurrentText(
            opts.epub_opt.reader_preset if opts.epub_opt.reader_preset in READER_PRESETS else "기본(일반)"
        )

        self.preset_desc = QLabel()
        self.preset_desc.setStyleSheet("color:#666;")

        form_preset.addRow("프리셋 선택", self.preset_combo)
        form_preset.addRow("설명", self.preset_desc)

        v_preset.addWidget(box_preset)
        v_preset.addStretch()
        tabs.addTab(page_preset, "프리셋")

        # ==========================================================
        # 5) 최적화(1) 탭 : 모드/서브셋/표지/ZIP (스크롤)
        # ==========================================================
        page_opt1 = QWidget()
        v_opt1 = QVBoxLayout(page_opt1)
        v_opt1.setSpacing(10)

        box_mode = QGroupBox("최적화 모드")
        form_mode = QFormLayout(box_mode)
        self.optimize_mode = QCheckBox("용량 최적화 모드(EPUB) 활성화")
        self.optimize_mode.setChecked(bool(opts.epub_opt.optimize_mode))
        form_mode.addRow("", self.optimize_mode)
        v_opt1.addWidget(box_mode)

        self.box_subset = QGroupBox("폰트 서브셋")
        form_subset = QFormLayout(self.box_subset)
        self.subset_font = QCheckBox("임베드 폰트 서브셋(사용 글자만 포함)")
        self.subset_font.setChecked(bool(opts.epub_opt.subset_font))
        self.prefer_internal = QCheckBox("내부 서브셋(fontTools 가능 시) 우선")
        self.prefer_internal.setChecked(bool(opts.epub_opt.prefer_internal_subsetter))
        form_subset.addRow("", self.subset_font)
        form_subset.addRow("", self.prefer_internal)
        v_opt1.addWidget(self.box_subset)

        self.box_cover = QGroupBox("표지 최적화")
        form_cover = QFormLayout(self.box_cover)
        self.cover_max_px = QSpinBox()
        self.cover_max_px.setRange(0, 3000)
        self.cover_max_px.setValue(int(opts.epub_opt.cover_max_px))
        self.cover_format = QComboBox()
        self.cover_format.addItems(["JPEG", "PNG"])
        self.cover_format.setCurrentText(opts.epub_opt.cover_format if opts.epub_opt.cover_format in ("JPEG", "PNG") else "JPEG")
        self.cover_jpeg_quality = QSpinBox()
        self.cover_jpeg_quality.setRange(40, 95)
        self.cover_jpeg_quality.setValue(int(opts.epub_opt.cover_jpeg_quality))
        form_cover.addRow("표지 최대 크기(px, 긴 변 기준, 0=원본)", self.cover_max_px)
        form_cover.addRow("표지 포맷", self.cover_format)
        form_cover.addRow("JPEG 품질", self.cover_jpeg_quality)
        v_opt1.addWidget(self.box_cover)

        self.box_zip = QGroupBox("ZIP 재압축")
        form_zip = QFormLayout(self.box_zip)
        self.repack_zip = QCheckBox("저장 후 ZIP 재압축(권장)")
        self.repack_zip.setChecked(bool(opts.epub_opt.repack_zip))
        self.zip_level = QSpinBox()
        self.zip_level.setRange(1, 9)
        self.zip_level.setValue(int(opts.epub_opt.zip_compresslevel))
        form_zip.addRow("", self.repack_zip)
        form_zip.addRow("압축 레벨(1~9)", self.zip_level)
        v_opt1.addWidget(self.box_zip)

        v_opt1.addStretch()
        scroll1 = QScrollArea()
        scroll1.setWidgetResizable(True)
        scroll1.setFrameShape(QFrame.NoFrame)
        scroll1.setWidget(page_opt1)
        tabs.addTab(scroll1, "최적화(1)")

        # ==========================================================
        # 6) 최적화(2) 탭 : 외부 도구 (스크롤)
        # ==========================================================
        page_opt2 = QWidget()
        v_opt2 = QVBoxLayout(page_opt2)
        v_opt2.setSpacing(10)

        self.box_ext_epub = QGroupBox("외부 EPUB 최적화 도구(선택)")
        form_ext_epub = QFormLayout(self.box_ext_epub)
        self.external_epub_enabled = QCheckBox("외부 도구 실행")
        self.external_epub_enabled.setChecked(bool(opts.epub_opt.external_epub_tool_enabled))
        self.external_epub_cmd = QLineEdit(opts.epub_opt.external_epub_tool_cmd)
        self.external_epub_cmd.setPlaceholderText('예) tool "{in}" "{out}"')
        form_ext_epub.addRow("", self.external_epub_enabled)
        form_ext_epub.addRow("커맨드 템플릿", self.external_epub_cmd)
        v_opt2.addWidget(self.box_ext_epub)

        self.box_ext_font = QGroupBox("외부 폰트 서브셋 도구(선택)")
        form_ext_font = QFormLayout(self.box_ext_font)
        self.external_font_enabled = QCheckBox("내부 서브셋 실패 시 외부 서브셋 사용")
        self.external_font_enabled.setChecked(bool(opts.epub_opt.external_font_subset_enabled))
        self.external_font_cmd = QLineEdit(opts.epub_opt.external_font_subset_cmd)
        self.external_font_cmd.setPlaceholderText('예) pyftsubset "{in}" --output-file="{out}" --text-file="{textfile}"')
        form_ext_font.addRow("", self.external_font_enabled)
        form_ext_font.addRow("커맨드 템플릿", self.external_font_cmd)
        v_opt2.addWidget(self.box_ext_font)

        v_opt2.addStretch()
        scroll2 = QScrollArea()
        scroll2.setWidgetResizable(True)
        scroll2.setFrameShape(QFrame.NoFrame)
        scroll2.setWidget(page_opt2)
        tabs.addTab(scroll2, "최적화(2)")

        # ==========================================================
        # 연결: 미리보기/프리셋/최적화 enable
        # ==========================================================
        self.pdf_font_combo.currentFontChanged.connect(self._refresh_previews)
        self.pdf_body_size.valueChanged.connect(self._refresh_previews)
        self.epub_font_combo.currentFontChanged.connect(self._refresh_previews)
        self.epub_base_em.valueChanged.connect(self._refresh_previews)

        self.preset_combo.currentTextChanged.connect(self._on_preset_changed)
        self.optimize_mode.toggled.connect(self._sync_optimize_enabled)

        # 초기 반영
        self._on_preset_changed(self.preset_combo.currentText())
        self._sync_optimize_enabled(self.optimize_mode.isChecked())
        self._refresh_previews()

        # OK/Cancel
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _sync_custom_paper_ui(self, name: str):
        is_custom = (name == "사용자 지정(mm)")
        self.custom_w.setEnabled(is_custom)
        self.custom_h.setEnabled(is_custom)

    def _sync_optimize_enabled(self, enabled: bool):
        for w in (self.box_subset, self.box_cover, self.box_zip, self.box_ext_epub, self.box_ext_font):
            w.setEnabled(bool(enabled))

    def _update_preset_desc(self, preset_name: str):
        p = READER_PRESETS.get(preset_name, {})
        if not p:
            self.preset_desc.setText("")
            return
        embed = "임베드" if p.get("embed_font") else "비임베드"
        cover = f"{p.get('cover_max_px', 1200)}px / {p.get('cover_format','JPEG')} Q{p.get('cover_jpeg_quality',82)}"
        zip_lvl = p.get("zip_level", 9)
        self.preset_desc.setText(f"- 폰트: {embed}\n- 표지: {cover}\n- ZIP: 재압축 / 레벨 {zip_lvl}")

    def _apply_preset_to_ui(self, preset_name: str):
        p = READER_PRESETS.get(preset_name)
        if not p:
            return
        self.epub_embed.setChecked(bool(p.get("embed_font", False)))
        self.cover_max_px.setValue(int(p.get("cover_max_px", 1200)))
        self.cover_format.setCurrentText(str(p.get("cover_format", "JPEG")))
        self.cover_jpeg_quality.setValue(int(p.get("cover_jpeg_quality", 82)))
        self.repack_zip.setChecked(bool(p.get("repack_zip", True)))
        self.zip_level.setValue(int(p.get("zip_level", 9)))

    def _on_preset_changed(self, preset_name: str):
        self._update_preset_desc(preset_name)
        self._apply_preset_to_ui(preset_name)

    def _refresh_previews(self):
        f_pdf = QFont(self.pdf_font_combo.currentFont().family())
        f_pdf.setPointSize(int(self.pdf_body_size.value()))
        self.pdf_preview.setFont(f_pdf)

        size_pt = int(round(12 * float(self.epub_base_em.value())))
        f_epub = QFont(self.epub_font_combo.currentFont().family())
        f_epub.setPointSize(max(8, size_pt))
        self.epub_preview.setFont(f_epub)

    def _pick_font_file(self, settings_key: str, caption: str, on_done):
        last_dir = self.settings.value(settings_key, "", type=str)
        if not last_dir:
            win_dir = str(Path(os.environ.get("SystemRoot", r"C:\Windows")) / "Fonts")
            last_dir = win_dir if os.path.exists(win_dir) else str(Path.home())

        self.dialogs.open_existing_file(
            caption=caption,
            name_filter="Font (*.ttf *.otf *.ttc *.otc)",
            last_dir_key=settings_key,
            on_done=on_done,
            directory=last_dir,
        )

    def _browse_pdf_font(self):
        def _done(path: str):
            if path:
                self.pdf_font_path.setText(path)
        self._pick_font_file("ui/last_dir_pdf_font", "PDF 폰트 파일 선택", _done)

    def _browse_epub_font(self):
        def _done(path: str):
            if path:
                self.epub_font_path.setText(path)
        self._pick_font_file("ui/last_dir_epub_font", "EPUB 폰트 파일 선택", _done)

    def get_options(self) -> TypesetOptions:
        pdf_prof = PdfFontProfile(
            family=self.pdf_font_combo.currentFont().family(),
            file_path=self.pdf_font_path.text().strip(),
            body_size_pt=int(self.pdf_body_size.value()),
            chapter_title_pt=int(self.pdf_chapter_size.value()),
            cover_title_pt=int(self.pdf_cover_size.value()),
        )

        epub_prof = EpubFontProfile(
            family=self.epub_font_combo.currentFont().family(),
            file_path=self.epub_font_path.text().strip(),
            base_font_em=float(self.epub_base_em.value()),
            embed_font=bool(self.epub_embed.isChecked()),
        )

        preview_profile = "pdf" if self.preview_profile_combo.currentIndex() == 0 else "epub"

        epub_opt = EpubOptimizeOptions(
            optimize_mode=bool(self.optimize_mode.isChecked()),
            reader_preset=self.preset_combo.currentText(),
            auto_typo_tune=bool(self.epub_auto_tune.isChecked()),
            subset_font=bool(self.subset_font.isChecked()),
            prefer_internal_subsetter=bool(self.prefer_internal.isChecked()),
            cover_max_px=int(self.cover_max_px.value()),
            cover_format=self.cover_format.currentText(),
            cover_jpeg_quality=int(self.cover_jpeg_quality.value()),
            repack_zip=bool(self.repack_zip.isChecked()),
            zip_compresslevel=int(self.zip_level.value()),
            external_epub_tool_enabled=bool(self.external_epub_enabled.isChecked()),
            external_epub_tool_cmd=self.external_epub_cmd.text().strip(),
            external_font_subset_enabled=bool(self.external_font_enabled.isChecked()),
            external_font_subset_cmd=self.external_font_cmd.text().strip(),
        )

        return TypesetOptions(
            paper=self.paper_combo.currentText(),
            custom_paper_w_mm=int(self.custom_w.value()),
            custom_paper_h_mm=int(self.custom_h.value()),
            line_spacing=float(self.line_spacing.value()),
            para_spacing=int(self.para_spacing.value()),
            indent_spaces=int(self.indent_spaces.value()),
            show_header=bool(self.header_check.isChecked()),
            include_title_page=bool(self.cover_check.isChecked()),
            chapter_strict=bool(self.strict_check.isChecked()),
            chapter_min_lines=int(self.min_lines.value()),
            preview_profile=preview_profile,
            pdf_profile=pdf_prof,
            epub_profile=epub_prof,
            epub_opt=epub_opt,
        )
        
# ==========================================================
# Part 3/3
# - Part1 Hotfix (필수 교체)
# - PDF/EPUB 저장 + EPUB CSS(판형/타이포/대사블록 자동)
# - 폰트 탐색/서브셋/표지/ZIP 재압축/용량리포트
# - MainWindow + 실행부
# ==========================================================
# ==========================================================
# 라이브러리 체크(에러 메시지 강화)
# ==========================================================
if FPDF is None:
    raise RuntimeError(
        "PDF 저장을 위해 fpdf2가 필요합니다.\n"
        "pip install --upgrade fpdf2"
    )


# ==========================================================
# 리더 프리셋 (EPUB 최적화 저장 시 반영)
# ==========================================================
READER_PRESETS: Dict[str, dict] = {
    "기본(일반)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=82, repack_zip=True, zip_level=9),
    "Kindle(추천)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=80, repack_zip=True, zip_level=9),
    "리디북스(추천)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=82, repack_zip=True, zip_level=9),
    "교보 eBook(추천)": dict(embed_font=False, cover_max_px=1200, cover_format="JPEG", cover_jpeg_quality=82, repack_zip=True, zip_level=9),
    "Apple Books(추천)": dict(embed_font=False, cover_max_px=1400, cover_format="JPEG", cover_jpeg_quality=85, repack_zip=True, zip_level=9),
    "폰트임베드(용량↑/호환성↑)": dict(embed_font=True, cover_max_px=1400, cover_format="JPEG", cover_jpeg_quality=85, repack_zip=True, zip_level=9),
}


# ==========================================================
# Async QFileDialog Helper (open() 기반)
# ==========================================================
class AsyncFileDialogs(QObject):
    def __init__(self, parent: QWidget, settings: QSettings):
        super().__init__(parent)
        self._parent = parent
        self._settings = settings
        self._alive: List[QFileDialog] = []

    def _start_dir(self, key: str) -> str:
        d = self._settings.value(key, "", type=str)
        return d if d else str(Path.home())

    def _remember_dir(self, key: str, file_path: str):
        try:
            self._settings.setValue(key, str(Path(file_path).parent))
        except Exception:
            pass

    def open_existing_file(self, caption: str, name_filter: str, last_dir_key: str, on_done, directory: Optional[str] = None):
        dlg = QFileDialog(self._parent, caption)
        dlg.setFileMode(QFileDialog.ExistingFile)
        dlg.setNameFilter(name_filter)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setDirectory(directory or self._start_dir(last_dir_key))
        self._alive.append(dlg)

        def _finished(result: int):
            try:
                if result == QDialog.Accepted:
                    files = dlg.selectedFiles()
                    chosen = files[0] if files else ""
                    if chosen:
                        self._remember_dir(last_dir_key, chosen)
                    on_done(chosen)
                else:
                    on_done("")
            finally:
                if dlg in self._alive:
                    self._alive.remove(dlg)
                dlg.deleteLater()

        dlg.finished.connect(_finished)
        dlg.open()

    def open_save_file(self, caption: str, default_name: str, name_filter: str, default_suffix: str,
                       last_dir_key: str, on_done, labels_ko: bool = True, directory: Optional[str] = None):
        dlg = QFileDialog(self._parent, caption)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilter(name_filter)
        dlg.setDefaultSuffix(default_suffix)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setDirectory(directory or self._start_dir(last_dir_key))
        dlg.selectFile(default_name)

        if labels_ko:
            dlg.setLabelText(QFileDialog.LookIn, "저장 위치:")
            dlg.setLabelText(QFileDialog.FileName, "파일 이름:")
            dlg.setLabelText(QFileDialog.FileType, "파일 형식:")
            dlg.setLabelText(QFileDialog.Accept, "저장")
            dlg.setLabelText(QFileDialog.Reject, "취소")

        self._alive.append(dlg)

        def _finished(result: int):
            try:
                if result == QDialog.Accepted:
                    files = dlg.selectedFiles()
                    chosen = files[0] if files else ""
                    if chosen:
                        self._remember_dir(last_dir_key, chosen)
                    on_done(chosen)
                else:
                    on_done("")
            finally:
                if dlg in self._alive:
                    self._alive.remove(dlg)
                dlg.deleteLater()

        dlg.finished.connect(_finished)
        dlg.open()


# ==========================================================
# Drag & Drop Text Editor
# ==========================================================
class DropTextEdit(QPlainTextEdit):
    files_dropped = Signal(list)  # List[str]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setPlaceholderText(
            "본문을 입력하거나\n"
            "텍스트 파일(.txt, .md)을 여기로 드래그하세요.\n"
            "여러 파일을 드롭하면 자동 합치기가 가능합니다."
        )

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for u in event.mimeData().urls():
                if u.isLocalFile():
                    p = u.toLocalFile().lower()
                    if p.endswith((".txt", ".md")):
                        event.acceptProposedAction()
                        return
        event.ignore()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        paths = []
        for u in event.mimeData().urls():
            if u.isLocalFile():
                p = u.toLocalFile()
                if p.lower().endswith((".txt", ".md")):
                    paths.append(p)
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()


# ==========================================================
# Utils
# ==========================================================
def contains_korean(text: str) -> bool:
    for ch in text:
        o = ord(ch)
        if (0xAC00 <= o <= 0xD7A3) or (0x1100 <= o <= 0x11FF) or (0x3130 <= o <= 0x318F):
            return True
    return False


# ==========================================================
# PDF 클래스
# ==========================================================
class NovelPDF(FPDF):
    def __init__(self, book_title="", show_header=True, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.book_title = book_title
        self.show_header = show_header
        self._body_family = "Helvetica"

    def set_body_font(self, family: str):
        self._body_family = family

    def header(self):
        if not self.show_header or not self.book_title:
            return
        self.set_font(self._body_family, size=8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 8, text=self.book_title, align="C")
        self.ln(10)

    def footer(self):
        self.set_y(-12)
        self.set_font(self._body_family, size=8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 10, text=str(self.page_no()), align="C")


# ==========================================================
# 비동기 추출 워커
# ==========================================================
class ExtractWorker(QObject):
    finished = Signal(str)
    error = Signal(str)
    log_signal = Signal(str)

    def __init__(self, url: str):
        super().__init__()
        self.url = url

    @Slot()
    def run(self):
        if not trafilatura:
            self.error.emit("trafilatura 라이브러리가 설치되지 않았습니다. (pip install trafilatura)")
            return
        try:
            self.log_signal.emit(f"데이터 추출 중: {self.url}")
            downloaded = trafilatura.fetch_url(self.url)
            if not downloaded:
                self.error.emit("URL 접근 실패 (네트워크 또는 사이트 차단)")
                return
            result = trafilatura.extract(downloaded, include_comments=False, include_links=False)
            if result:
                self.finished.emit(result)
            else:
                self.error.emit("본문을 추출할 수 없습니다.")
        except Exception as e:
            self.error.emit(str(e))


# ==========================================================
# MainWindow
# ==========================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("MySoft", "Web2Multi_PRO_Final_AllInOne")
        self.setWindowTitle("Web2Multi PRO - Final All-in-One (2026.04)")
        self.resize(1500, 940)

        self.temp_dir = Path(tempfile.mkdtemp(prefix="w2m_final_"))
        self.windows_fonts_dir = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "Fonts"

        self.thread: Optional[QThread] = None
        self.worker: Optional[ExtractWorker] = None

        self.ui_font_size = int(self.settings.value("ui/font_size", 12))
        self.opts = self.load_typeset_options()
        self.dialogs = AsyncFileDialogs(self, self.settings)

        self.create_actions()
        self.init_ui()
        self.init_menus()

        self.apply_ui_font_size(self.ui_font_size)
        self.apply_styles()
        self.restore_ui_state()
        self.update_preview()

    # ------------------------
    # Menu font size sync
    # ------------------------
    def apply_menu_font_size(self, pt: int):
        qss = f"""
        QMenuBar {{
            font-size: {pt}pt;
            padding: 4px;
        }}
        QMenuBar::item {{
            spacing: 6px;
            padding: 4px 10px;
        }}
        QMenu {{
            font-size: {pt}pt;
            padding: 6px;
        }}
        QMenu::item {{
            padding: 6px 18px;
        }}
        QMenu::item:selected {{
            background-color: rgba(80,160,255,0.25);
        }}
        """
        self.menuBar().setStyleSheet(qss)

    # ------------------------
    # Icons / Log
    # ------------------------
    def icon(self, theme_name: str, fallback_sp: QStyle.StandardPixmap) -> QIcon:
        ico = QIcon.fromTheme(theme_name)
        if not ico.isNull():
            return ico
        return self.style().standardIcon(fallback_sp)

    def log(self, msg: str):
        self.status_bar.showMessage(f"상태: {msg}")

    # ------------------------
    # Paper format mm
    # ------------------------
    def paper_format_mm(self) -> Tuple[int, int]:
        name = self.opts.paper
        if name == "사용자 지정(mm)":
            return (int(self.opts.custom_paper_w_mm), int(self.opts.custom_paper_h_mm))
        wh = BOOK_TRIM_SIZES_MM.get(name)
        if wh:
            return wh
        return (148, 210)

    # ==========================================================
    # Typeset options load/save (하위호환 포함)
    # ==========================================================
    def load_typeset_options(self) -> TypesetOptions:
        # PDF profile
        pdf_prof = PdfFontProfile(
            family=self.settings.value("pdf/font_family", "맑은 고딕", type=str),
            file_path=self.settings.value("pdf/font_path", "", type=str),
            body_size_pt=int(self.settings.value("pdf/body_pt", 11)),
            chapter_title_pt=int(self.settings.value("pdf/chapter_pt", 15)),
            cover_title_pt=int(self.settings.value("pdf/cover_pt", 24)),
        )
        # EPUB profile
        epub_prof = EpubFontProfile(
            family=self.settings.value("epub/font_family", "맑은 고딕", type=str),
            file_path=self.settings.value("epub/font_path", "", type=str),
            base_font_em=float(self.settings.value("epub/base_em", 1.0)),
            embed_font=self.settings.value("epub/embed_font", False, type=bool),
        )

        saved_paper = self.settings.value("typeset/paper", "국판(A5) 148×210", type=str)
        # 하위호환: 예전 값
        if saved_paper == "A5":
            saved_paper = "국판(A5) 148×210"
        elif saved_paper == "A4":
            saved_paper = "A4 210×297"
        if saved_paper not in BOOK_TRIM_SIZES_MM:
            saved_paper = "국판(A5) 148×210"

        epub_opt = EpubOptimizeOptions(
            optimize_mode=self.settings.value("epubopt/optimize_mode", True, type=bool),
            reader_preset=self.settings.value("epubopt/reader_preset", "기본(일반)", type=str),
            auto_typo_tune=self.settings.value("epubopt/auto_typo_tune", True, type=bool),
            subset_font=self.settings.value("epubopt/subset_font", True, type=bool),
            prefer_internal_subsetter=self.settings.value("epubopt/prefer_internal", True, type=bool),
            cover_max_px=int(self.settings.value("epubopt/cover_max_px", 1200)),
            cover_format=self.settings.value("epubopt/cover_format", "JPEG", type=str),
            cover_jpeg_quality=int(self.settings.value("epubopt/cover_jpeg_quality", 82)),
            repack_zip=self.settings.value("epubopt/repack_zip", True, type=bool),
            zip_compresslevel=int(self.settings.value("epubopt/zip_level", 9)),
            external_epub_tool_enabled=self.settings.value("epubopt/external_epub_enabled", False, type=bool),
            external_epub_tool_cmd=self.settings.value("epubopt/external_epub_cmd", "", type=str),
            external_font_subset_enabled=self.settings.value("epubopt/external_font_enabled", False, type=bool),
            external_font_subset_cmd=self.settings.value("epubopt/external_font_cmd", "", type=str),
        )

        return TypesetOptions(
            paper=saved_paper,
            custom_paper_w_mm=int(self.settings.value("typeset/custom_w_mm", 148)),
            custom_paper_h_mm=int(self.settings.value("typeset/custom_h_mm", 210)),
            line_spacing=float(self.settings.value("typeset/line_spacing", 1.8)),
            para_spacing=int(self.settings.value("typeset/para_spacing", 2)),
            indent_spaces=int(self.settings.value("typeset/indent_spaces", 4)),
            show_header=self.settings.value("typeset/show_header", True, type=bool),
            include_title_page=self.settings.value("typeset/include_title_page", True, type=bool),
            chapter_strict=self.settings.value("typeset/chapter_strict", True, type=bool),
            chapter_min_lines=int(self.settings.value("typeset/chapter_min_lines", 3)),
            preview_profile=self.settings.value("typeset/preview_profile", "pdf", type=str),
            pdf_profile=pdf_prof,
            epub_profile=epub_prof,
            epub_opt=epub_opt,
        )

    def save_typeset_options(self, opts: TypesetOptions):
        self.settings.setValue("typeset/paper", opts.paper)
        self.settings.setValue("typeset/custom_w_mm", opts.custom_paper_w_mm)
        self.settings.setValue("typeset/custom_h_mm", opts.custom_paper_h_mm)
        self.settings.setValue("typeset/line_spacing", opts.line_spacing)
        self.settings.setValue("typeset/para_spacing", opts.para_spacing)
        self.settings.setValue("typeset/indent_spaces", opts.indent_spaces)
        self.settings.setValue("typeset/show_header", opts.show_header)
        self.settings.setValue("typeset/include_title_page", opts.include_title_page)
        self.settings.setValue("typeset/chapter_strict", opts.chapter_strict)
        self.settings.setValue("typeset/chapter_min_lines", opts.chapter_min_lines)
        self.settings.setValue("typeset/preview_profile", opts.preview_profile)

        self.settings.setValue("pdf/font_family", opts.pdf_profile.family)
        self.settings.setValue("pdf/font_path", opts.pdf_profile.file_path)
        self.settings.setValue("pdf/body_pt", opts.pdf_profile.body_size_pt)
        self.settings.setValue("pdf/chapter_pt", opts.pdf_profile.chapter_title_pt)
        self.settings.setValue("pdf/cover_pt", opts.pdf_profile.cover_title_pt)

        self.settings.setValue("epub/font_family", opts.epub_profile.family)
        self.settings.setValue("epub/font_path", opts.epub_profile.file_path)
        self.settings.setValue("epub/base_em", opts.epub_profile.base_font_em)
        self.settings.setValue("epub/embed_font", opts.epub_profile.embed_font)

        self.settings.setValue("epubopt/optimize_mode", opts.epub_opt.optimize_mode)
        self.settings.setValue("epubopt/reader_preset", opts.epub_opt.reader_preset)
        self.settings.setValue("epubopt/auto_typo_tune", opts.epub_opt.auto_typo_tune)
        self.settings.setValue("epubopt/subset_font", opts.epub_opt.subset_font)
        self.settings.setValue("epubopt/prefer_internal", opts.epub_opt.prefer_internal_subsetter)
        self.settings.setValue("epubopt/cover_max_px", opts.epub_opt.cover_max_px)
        self.settings.setValue("epubopt/cover_format", opts.epub_opt.cover_format)
        self.settings.setValue("epubopt/cover_jpeg_quality", opts.epub_opt.cover_jpeg_quality)
        self.settings.setValue("epubopt/repack_zip", opts.epub_opt.repack_zip)
        self.settings.setValue("epubopt/zip_level", opts.epub_opt.zip_compresslevel)
        self.settings.setValue("epubopt/external_epub_enabled", opts.epub_opt.external_epub_tool_enabled)
        self.settings.setValue("epubopt/external_epub_cmd", opts.epub_opt.external_epub_tool_cmd)
        self.settings.setValue("epubopt/external_font_enabled", opts.epub_opt.external_font_subset_enabled)
        self.settings.setValue("epubopt/external_font_cmd", opts.epub_opt.external_font_subset_cmd)

    # ==========================================================
    # Actions
    # ==========================================================
    def create_actions(self):
        self.act_open_text = QAction(self.icon("document-open", QStyle.SP_DialogOpenButton), "Text 파일 불러오기...", self)
        self.act_open_text.setShortcut(QKeySequence("Ctrl+O"))
        self.act_open_text.triggered.connect(self.open_text_files)

        self.act_extract = QAction(self.icon("download", QStyle.SP_ArrowDown), "본문 추출", self)
        self.act_extract.setShortcut(QKeySequence("Ctrl+E"))
        self.act_extract.triggered.connect(self.on_start_extraction)

        self.act_save = QAction(self.icon("document-save", QStyle.SP_DialogSaveButton), "저장", self)
        self.act_save.setShortcut(QKeySequence("Ctrl+S"))
        self.act_save.triggered.connect(self.on_save_clicked)
        self.act_save.setEnabled(False)

        self.act_reset = QAction(self.icon("edit-clear", QStyle.SP_DialogResetButton), "전체 초기화", self)
        self.act_reset.setShortcut(QKeySequence("Ctrl+R"))
        self.act_reset.triggered.connect(self.reset_window)

        self.act_exit = QAction(self.icon("application-exit", QStyle.SP_DialogCloseButton), "종료", self)
        self.act_exit.setShortcut(QKeySequence("Ctrl+Q"))
        self.act_exit.triggered.connect(self.close)

        self.act_typeset_settings = QAction(self.icon("preferences-system", QStyle.SP_FileDialogDetailedView), "조판 설정...", self)
        self.act_typeset_settings.triggered.connect(self.open_typeset_dialog)

    # ==========================================================
    # UI
    # ==========================================================
    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        self.topbar = QFrame()
        self.topbar.setObjectName("TopBar")
        top = QHBoxLayout(self.topbar)
        top.setContentsMargins(12, 10, 12, 10)
        top.setSpacing(10)

        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("웹 소설 URL을 입력하세요...")
        self.url_edit.setText(self.settings.value("ui/last_url", "", type=str))

        btn_extract = QToolButton()
        btn_extract.setDefaultAction(self.act_extract)
        btn_extract.setAutoRaise(True)
        btn_extract.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)

        self.format_combo = QComboBox()
        self.format_combo.addItems(["PDF (소설 조판)", "EPUB (전자책)"])
        last_fmt = int(self.settings.value("ui/last_format", 0))
        if 0 <= last_fmt < self.format_combo.count():
            self.format_combo.setCurrentIndex(last_fmt)

        btn_save = QToolButton()
        btn_save.setDefaultAction(self.act_save)
        btn_save.setAutoRaise(True)
        btn_save.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)

        btn_reset = QToolButton()
        btn_reset.setDefaultAction(self.act_reset)
        btn_reset.setAutoRaise(True)
        btn_reset.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)

        self.title_edit = QLineEdit(self.settings.value("meta/title", "제목 없음"))
        self.author_edit = QLineEdit(self.settings.value("meta/author", "작자 미상"))
        self.title_edit.setPlaceholderText("제목")
        self.author_edit.setPlaceholderText("저자")

        top.addWidget(QLabel("URL:"))
        top.addWidget(self.url_edit, 3)
        top.addWidget(btn_extract)
        top.addSpacing(8)
        top.addWidget(QLabel("제목:"))
        top.addWidget(self.title_edit, 2)
        top.addWidget(QLabel("저자:"))
        top.addWidget(self.author_edit, 1)
        top.addSpacing(8)
        top.addWidget(QLabel("출력:"))
        top.addWidget(self.format_combo, 1)
        top.addWidget(btn_save)
        top.addWidget(btn_reset)

        root.addWidget(self.topbar)

        self.splitter = QSplitter(Qt.Horizontal)
        self.editor = DropTextEdit()
        self.editor.textChanged.connect(self.on_text_changed)
        self.editor.files_dropped.connect(self.on_files_dropped)

        self.preview_view = QTextBrowser()
        self.splitter.addWidget(self.editor)
        self.splitter.addWidget(self.preview_view)
        self.splitter.setSizes([750, 750])
        root.addWidget(self.splitter, 1)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.title_edit.textChanged.connect(self.update_preview)
        self.author_edit.textChanged.connect(self.update_preview)

    def init_menus(self):
        mbar = self.menuBar()

        file_menu = mbar.addMenu("파일(&F)")
        file_menu.addAction(self.act_open_text)
        file_menu.addSeparator()
        file_menu.addAction(self.act_extract)
        file_menu.addAction(self.act_save)
        file_menu.addSeparator()
        file_menu.addAction(self.act_reset)
        file_menu.addSeparator()
        file_menu.addAction(self.act_exit)

        typeset_menu = mbar.addMenu("조판(&T)")
        typeset_menu.addAction(self.act_typeset_settings)

        setup_menu = mbar.addMenu("설정(&S)")
        font_size_menu = setup_menu.addMenu("UI 글자 크기")

        self.font_group = QActionGroup(self)
        self.font_group.setExclusive(True)

        for size in range(10, 26, 2):
            act = QAction(f"{size}pt", self, checkable=True)
            act.setData(size)
            if size == self.ui_font_size:
                act.setChecked(True)
            act.triggered.connect(self._on_menu_ui_font_changed)
            self.font_group.addAction(act)
            font_size_menu.addAction(act)

    def _on_menu_ui_font_changed(self):
        act = self.sender()
        if not act:
            return
        self.apply_ui_font_size(int(act.data()))

    def apply_styles(self):
        self.setStyleSheet("""
            #TopBar {
                border: 1px solid rgba(120,120,120,0.28);
                border-radius: 14px;
                background: rgba(120,120,120,0.06);
            }
            QLineEdit, QComboBox {
                padding: 7px 10px;
                border-radius: 12px;
                border: 1px solid rgba(120,120,120,0.25);
            }
            QLineEdit:focus, QComboBox:focus {
                border: 1px solid rgba(80,160,255,0.65);
            }
            QToolButton {
                padding: 8px 10px;
                border-radius: 12px;
            }
            QToolButton:hover { background: rgba(120,120,120,0.12); }
            QToolButton:pressed { background: rgba(120,120,120,0.22); }
            QPlainTextEdit, QTextBrowser {
                border-radius: 16px;
                border: 1px solid rgba(120,120,120,0.35);
            }
        """)

    def apply_ui_font_size(self, size: int):
        self.ui_font_size = int(size)
        self.settings.setValue("ui/font_size", self.ui_font_size)
        f = QFontDatabase.systemFont(QFontDatabase.GeneralFont)
        f.setPointSize(self.ui_font_size)
        QApplication.setFont(f)
        self.apply_menu_font_size(self.ui_font_size)

        if hasattr(self, "font_group"):
            for act in self.font_group.actions():
                if int(act.data()) == self.ui_font_size:
                    act.setChecked(True)
        self.update_preview()

    # ==========================================================
    # 조판 설정 다이얼로그 (Part 2/3에 정의된 TypesetDialog 사용)
    # ==========================================================
    def open_typeset_dialog(self):
        # Part 2/3에 TypesetDialog가 있어야 합니다.
        dlg = TypesetDialog(self, self.opts, self.settings)
        if dlg.exec() == QDialog.Accepted:
            self.opts = dlg.get_options()
            self.save_typeset_options(self.opts)
            self.update_preview()
            self.log("조판 설정이 저장되었습니다.")

    # ==========================================================
    # Text change / file open / drag drop
    # ==========================================================
    def on_text_changed(self):
        has_text = bool(self.editor.toPlainText().strip())
        self.act_save.setEnabled(has_text)
        self.update_preview()

    def _read_text_file(self, path: str) -> str:
        for enc in ("utf-8-sig", "utf-8", "cp949"):
            try:
                with open(path, "r", encoding=enc) as f:
                    return f.read()
            except Exception:
                continue
        raise RuntimeError("지원 인코딩(UTF-8, CP949)으로 읽을 수 없습니다.")

    def _ask_drop_mode(self, nfiles: int) -> str:
        msg = QMessageBox(self)
        msg.setWindowTitle("적용 방식 선택")
        msg.setText(f"{nfiles}개 파일을 불러옵니다.\n어떻게 적용할까요?")
        btn_over = msg.addButton("덮어쓰기", QMessageBox.AcceptRole)
        btn_app = msg.addButton("이어붙이기", QMessageBox.AcceptRole)
        btn_cancel = msg.addButton("취소", QMessageBox.RejectRole)
        msg.setDefaultButton(btn_over)
        msg.exec()
        clicked = msg.clickedButton()
        if clicked == btn_cancel:
            return "cancel"
        if clicked == btn_app:
            return "append"
        return "overwrite"

    def open_text_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Text 파일 불러오기",
            self.settings.value("ui/last_dir_text", str(Path.home()), type=str),
            "Text Files (*.txt *.md);;All Files (*)"
        )
        if not paths:
            return
        self.settings.setValue("ui/last_dir_text", str(Path(paths[0]).parent))

        texts = []
        errors = []
        for p in paths:
            try:
                texts.append(self._read_text_file(p))
            except Exception as e:
                errors.append(f"{p}\n  - {e}")

        if errors:
            QMessageBox.warning(self, "일부 파일 로드 실패", "\n\n".join(errors))
        if not texts:
            return

        merged = "\n\n".join(texts) if len(texts) > 1 else texts[0]
        current = self.editor.toPlainText()

        mode = self._ask_drop_mode(len(paths))
        if mode == "cancel":
            return

        if mode == "overwrite":
            self.editor.setPlainText(merged)
        else:
            if current.strip():
                self.editor.setPlainText(current.rstrip() + "\n\n" + merged.lstrip())
            else:
                self.editor.setPlainText(merged)

    def on_files_dropped(self, paths: List[str]):
        if not paths:
            return
        texts = []
        errors = []
        for p in paths:
            try:
                texts.append(self._read_text_file(p))
            except Exception as e:
                errors.append(f"{p}\n  - {e}")
        if errors:
            QMessageBox.warning(self, "일부 파일 로드 실패", "\n\n".join(errors))
        if not texts:
            return

        merged = "\n\n".join(texts) if len(texts) > 1 else texts[0]
        current = self.editor.toPlainText()

        mode = self._ask_drop_mode(len(paths))
        if mode == "cancel":
            return
        if mode == "overwrite":
            self.editor.setPlainText(merged)
        else:
            if current.strip():
                self.editor.setPlainText(current.rstrip() + "\n\n" + merged.lstrip())
            else:
                self.editor.setPlainText(merged)

    # ==========================================================
    # Reset
    # ==========================================================
    def reset_window(self):
        reply = QMessageBox.question(
            self, "전체 초기화",
            "현재 편집 중인 내용이 모두 삭제됩니다.\n계속하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        if self.thread is not None:
            try:
                self.thread.quit()
                self.thread.wait(100)
            except Exception:
                pass
            self.thread = None
            self.worker = None

        self.editor.clear()
        self.preview_view.clear()
        self.title_edit.setText("제목 없음")
        self.author_edit.setText("작자 미상")

        self.settings.setValue("meta/title", "제목 없음")
        self.settings.setValue("meta/author", "작자 미상")

        self.act_extract.setEnabled(True)
        self.act_save.setEnabled(False)
        self.log("초기화 완료 — 새로운 본문 추출 가능")

    # ==========================================================
    # Preview
    # ==========================================================
    @Slot()
    def update_preview(self):
        content = self.editor.toPlainText()
        if not content.strip():
            self.preview_view.setHtml(
                "<div style='color:#888; padding:40px; text-align:center;'>"
                "좌측에 본문을 입력하거나 URL에서 추출해주세요.<br/>"
                "또는 TXT/MD 파일을 드래그앤드롭하세요."
                "</div>"
            )
            return

        chapters = split_into_chapters(content)
        # preview font from profile
        if self.opts.preview_profile == "pdf":
            prev_family = self.opts.pdf_profile.family
            prev_size_pt = self.opts.pdf_profile.body_size_pt
        else:
            prev_family = self.opts.epub_profile.family
            prev_size_pt = int(round(12 * self.opts.epub_profile.base_font_em))

        w_mm, h_mm = self.paper_format_mm()
        paper_info = f"{escape(self.opts.paper)} ({w_mm}×{h_mm}mm)"

        body_parts = [f"<div style='color:#777; font-size:10pt; margin-bottom:10px;'>판형: {paper_info}</div>"]
        for ch in chapters:
            body_parts.append(
                f"<h2 style='text-align:center; margin:0 0 14px 0;'>{escape(ch['title'])}</h2>"
            )
            for p in smart_make_paragraphs(ch["lines"]):
                if not p.strip():
                    body_parts.append("<div style='height:10px;'></div>")
                    continue
                stripped = p.lstrip()
                is_dlg = stripped.startswith(DIALOGUE_PREFIXES)
                indent = "0" if is_dlg else "1.5em"
                body_parts.append(
                    f"<p style='text-indent:{indent}; line-height:{self.opts.line_spacing}; "
                    f"margin:0 0 {self.opts.para_spacing}px 0;'>{escape(stripped)}</p>"
                )

        html = (
            "<style>"
            f"body{{background:#fff; color:#1a1a1a; font-size:{prev_size_pt}pt; padding:20px;"
            f"font-family:'{escape(prev_family)}','Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif;}}"
            "h2{font-weight:900; letter-spacing:-0.3px;}"
            "</style>"
            "<body>" + "".join(body_parts) + "</body>"
        )
        self.preview_view.setHtml(html)

    # ==========================================================
    # Extraction (thread)
    # ==========================================================
    def on_start_extraction(self):
        url = self.url_edit.text().strip()
        if not url:
            return

        if self.thread and self.thread.isRunning():
            self.log("이미 추출 작업이 진행 중입니다.")
            return

        self.act_extract.setEnabled(False)
        self.log("추출 중...")

        self.thread = QThread(self)
        self.worker = ExtractWorker(url)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log_signal.connect(self.log)

        self.worker.finished.connect(self._on_extract_success)
        self.worker.error.connect(self._on_extract_error)

        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def _cleanup_thread_refs(self):
        self.thread = None
        self.worker = None

    def _on_extract_success(self, text: str):
        self.editor.setPlainText(text)
        self.act_extract.setEnabled(True)
        self.log("추출 완료")
        self._cleanup_thread_refs()

    def _on_extract_error(self, err: str):
        QMessageBox.warning(self, "오류", err)
        self.act_extract.setEnabled(True)
        self.log("추출 실패")
        self._cleanup_thread_refs()

    # ==========================================================
    # Save
    # ==========================================================
    def on_save_clicked(self):
        content = self.editor.toPlainText().strip()
        if not content:
            return

        self.settings.setValue("meta/title", self.title_edit.text())
        self.settings.setValue("meta/author", self.author_edit.text())
        self.settings.setValue("ui/last_url", self.url_edit.text().strip())
        self.settings.setValue("ui/last_format", self.format_combo.currentIndex())

        is_pdf = "PDF" in self.format_combo.currentText()
        ext = "pdf" if is_pdf else "epub"
        default_name = f"result.{ext}"

        def _after_choose(path: str):
            if not path:
                return
            try:
                if is_pdf:
                    self.save_as_pdf(path, content)
                else:
                    self.save_as_epub(path, content)
                QMessageBox.information(self, "완료", "파일이 저장되었습니다.")
                self.log("저장 완료")
            except Exception:
                tb = traceback.format_exc()
                QMessageBox.critical(self, "저장 에러", tb)
                self.log("저장 실패")

        self.dialogs.open_save_file(
            caption="저장",
            default_name=default_name,
            name_filter=f"{ext.upper()} (*.{ext})",
            default_suffix=ext,
            last_dir_key="ui/last_dir_save",
            on_done=_after_choose,
            labels_ko=True,
        )

    # ==========================================================
    # Font resolving
    # ==========================================================
    def _font_search_dirs(self) -> List[Path]:
        dirs: List[Path] = []
        if self.windows_fonts_dir.exists():
            dirs.append(self.windows_fonts_dir)
        dirs.extend([Path("/System/Library/Fonts"), Path("/Library/Fonts")])
        dirs.extend([Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
                     Path.home() / ".fonts", Path.home() / ".local/share/fonts"])
        return [d for d in dirs if d.exists()]

    def _resolve_from_direct_path(self, path: str) -> Tuple[str, str]:
        if not path:
            return ("", "")
        p = Path(path)
        if p.exists():
            return (str(p), str(p))
        return ("", "")

    def resolve_font_file_from_family(self, family: str) -> Tuple[str, str]:
        if not family:
            return ("", "")

        win = self.windows_fonts_dir
        if family in ("맑은 고딕", "Malgun Gothic"):
            reg = win / "malgun.ttf"
            bold = win / "malgunbd.ttf"
            if reg.exists():
                return (str(reg), str(bold if bold.exists() else reg))

        target = re.sub(r"\s+", "", family).lower()

        candidates: List[Path] = []
        for d in self._font_search_dirs():
            candidates.extend(list(d.glob("*.ttf")))
            candidates.extend(list(d.glob("*.otf")))
            candidates.extend(list(d.glob("*.ttc")))
            candidates.extend(list(d.glob("*.otc")))

        matched = []
        for p in candidates:
            name = re.sub(r"\s+", "", p.stem).lower()
            if target and target in name:
                matched.append(p)

        if not matched:
            target2 = re.sub(r"[^a-z0-9가-힣]+", "", family.lower())
            for p in candidates:
                name2 = re.sub(r"[^a-z0-9가-힣]+", "", p.stem.lower())
                if target2 and target2 in name2:
                    matched.append(p)

        if not matched:
            return ("", "")

        bold_keys = ("bold", "bd", "demibold", "semibold")

        def score(p: Path) -> int:
            n = p.stem.lower()
            sc = 0
            if "regular" in n or "book" in n or "normal" in n:
                sc += 5
            if any(k in n for k in bold_keys):
                sc -= 2
            return sc

        matched.sort(key=score, reverse=True)
        regular = str(matched[0])

        bold = ""
        for p in matched:
            n = p.stem.lower()
            if any(k in n for k in bold_keys):
                bold = str(p)
                break
        if not bold:
            bold = regular

        return (regular, bold)

    def pdf_font_paths(self) -> Tuple[str, str]:
        reg, bold = self._resolve_from_direct_path(self.opts.pdf_profile.file_path)
        if reg:
            return (reg, bold)
        return self.resolve_font_file_from_family(self.opts.pdf_profile.family)

    def epub_font_path(self) -> str:
        reg, _ = self._resolve_from_direct_path(self.opts.epub_profile.file_path)
        if reg:
            return reg
        reg, _ = self.resolve_font_file_from_family(self.opts.epub_profile.family)
        return reg

    # ==========================================================
    # PDF save
    # ==========================================================
    def register_pdf_font(self, pdf: NovelPDF) -> str:
        reg_path, bold_path = self.pdf_font_paths()
        if not reg_path:
            return "Helvetica"

        dst_reg = self.temp_dir / Path(reg_path).name
        if not dst_reg.exists():
            shutil.copy2(reg_path, dst_reg)

        dst_bold = self.temp_dir / Path(bold_path).name
        if bold_path and not dst_bold.exists():
            shutil.copy2(bold_path, dst_bold)

        pdf_family = "NovelFont"
        pdf.add_font(family=pdf_family, style="", fname=str(dst_reg))
        pdf.add_font(family=pdf_family, style="B", fname=str(dst_bold if bold_path else dst_reg))
        return pdf_family

    def _pdf_add_title_page(self, pdf: NovelPDF, family: str):
        title = (self.title_edit.text().strip() or "제목 없음")
        author = (self.author_edit.text().strip() or "작자 미상")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        pdf.add_page()
        pdf.set_font(family, style="B", size=self.opts.pdf_profile.cover_title_pt)
        pdf.set_text_color(30, 30, 30)
        pdf.ln(40)
        pdf.multi_cell(0, 12, text=title, align="C")

        pdf.ln(12)
        pdf.set_font(family, size=max(10, self.opts.pdf_profile.body_size_pt + 3))
        pdf.set_text_color(80, 80, 80)
        pdf.cell(0, 10, text=f"저자: {author}", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.ln(6)
        pdf.set_font(family, size=max(8, self.opts.pdf_profile.body_size_pt - 1))
        pdf.set_text_color(120, 120, 120)
        pdf.cell(0, 8, text=f"생성일: {now_str}", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.ln(18)
        pdf.set_draw_color(160, 160, 160)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
        pdf.set_text_color(0, 0, 0)

    def save_as_pdf(self, path: str, content: str):
        if contains_korean(content):
            reg, _ = self.pdf_font_paths()
            if not reg:
                raise Exception(
                    "한글 본문이 감지되었지만, PDF 폰트 파일을 찾지 못했습니다.\n"
                    f"- PDF 폰트: {self.opts.pdf_profile.family}\n"
                    f"- PDF 폰트 파일: {self.opts.pdf_profile.file_path or '(미지정)'}\n\n"
                    "조판 설정 → PDF 프로파일에서 폰트 파일을 지정하세요."
                )

        page_format = self.paper_format_mm()

        pdf = NovelPDF(
            book_title=self.title_edit.text(),
            show_header=self.opts.show_header,
            format=page_format
        )
        pdf.set_auto_page_break(True, margin=15)

        family = self.register_pdf_font(pdf)
        pdf.set_body_font(family)
        pdf.set_font(family, size=self.opts.pdf_profile.body_size_pt)

        if self.opts.include_title_page:
            self._pdf_add_title_page(pdf, family)

        chapters = split_into_chapters(content)
        line_h = max(4.0, self.opts.line_spacing * (self.opts.pdf_profile.body_size_pt * 0.45))

        for ch in chapters:
            pdf.add_page()
            pdf.set_font(family, style="B", size=self.opts.pdf_profile.chapter_title_pt)
            pdf.cell(0, 15, text=ch["title"], align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(5)

            pdf.set_font(family, size=self.opts.pdf_profile.body_size_pt)
            for p in smart_make_paragraphs(ch["lines"]):
                if not p.strip():
                    pdf.ln(5)
                    continue
                txt = p if p.lstrip().startswith(DIALOGUE_PREFIXES) else (" " * self.opts.indent_spaces + p)
                pdf.multi_cell(0, line_h, text=txt, align="J")
                if self.opts.para_spacing:
                    pdf.ln(self.opts.para_spacing * 0.5)
                pdf.ln(5)  # 문단 분리를 위한 빈 줄 추가

        pdf.output(path)

    # ==========================================================
    # EPUB typography preset by trim
    # ==========================================================
    def _epub_typo_preset_by_trim(self, w_mm: int) -> dict:
        if w_mm <= 112:  # 문고판급
            return dict(
                indent_em=1.10, para_gap_em=0.70, dialog_gap_em=0.78,
                h1_mb_em=1.15, h1_ls_em=-0.020, dlg_side_pad_em=0.00,
                dlg_inblock_gap_em=0.28, dlg_block_top_em=0.55, dlg_block_bottom_em=0.78
            )
        elif w_mm <= 135:  # 46판/130변형
            return dict(
                indent_em=1.25, para_gap_em=0.66, dialog_gap_em=0.74,
                h1_mb_em=1.18, h1_ls_em=-0.018, dlg_side_pad_em=0.00,
                dlg_inblock_gap_em=0.26, dlg_block_top_em=0.52, dlg_block_bottom_em=0.74
            )
        elif w_mm <= 152:  # 국판~신국판
            return dict(
                indent_em=1.50, para_gap_em=0.60, dialog_gap_em=0.70,
                h1_mb_em=1.20, h1_ls_em=-0.016, dlg_side_pad_em=0.00,
                dlg_inblock_gap_em=0.24, dlg_block_top_em=0.50, dlg_block_bottom_em=0.70
            )
        elif w_mm <= 176:
            return dict(
                indent_em=1.55, para_gap_em=0.58, dialog_gap_em=0.68,
                h1_mb_em=1.22, h1_ls_em=-0.014, dlg_side_pad_em=0.00,
                dlg_inblock_gap_em=0.23, dlg_block_top_em=0.48, dlg_block_bottom_em=0.68
            )
        elif w_mm <= 190:
            return dict(
                indent_em=1.60, para_gap_em=0.55, dialog_gap_em=0.65,
                h1_mb_em=1.25, h1_ls_em=-0.012, dlg_side_pad_em=0.00,
                dlg_inblock_gap_em=0.22, dlg_block_top_em=0.46, dlg_block_bottom_em=0.66
            )
        else:
            return dict(
                indent_em=1.65, para_gap_em=0.52, dialog_gap_em=0.62,
                h1_mb_em=1.28, h1_ls_em=-0.010, dlg_side_pad_em=0.00,
                dlg_inblock_gap_em=0.20, dlg_block_top_em=0.44, dlg_block_bottom_em=0.64
            )

    def _epub_readability_css_by_trim(self, w_mm: int, base_em: float, auto_tune: bool, font_stack: str) -> str:
        # 판형별 판면(줄 길이/패딩)
        if w_mm <= 112:
            maxw_em, pad_em, font_scale = 32, 1.2, 1.05
        elif w_mm <= 135:
            maxw_em, pad_em, font_scale = 34, 1.3, 1.02
        elif w_mm <= 152:
            maxw_em, pad_em, font_scale = 38, 1.6, 1.00
        elif w_mm <= 176:
            maxw_em, pad_em, font_scale = 42, 1.9, 1.00
        elif w_mm <= 190:
            maxw_em, pad_em, font_scale = 46, 2.1, 0.98
        else:
            maxw_em, pad_em, font_scale = 52, 2.4, 0.98

        eff_em = max(0.7, min(1.6, base_em * font_scale))

        # 타이포 자동 튜닝
        if auto_tune:
            tp = self._epub_typo_preset_by_trim(w_mm)
            indent_em = tp["indent_em"]
            para_gap_em = tp["para_gap_em"]
            dialog_gap_em = tp["dialog_gap_em"]
            h1_mb_em = tp["h1_mb_em"]
            h1_ls_em = tp["h1_ls_em"]
            dlg_side_pad_em = tp.get("dlg_side_pad_em", 0.0)
            dlg_inblock_gap_em = tp.get("dlg_inblock_gap_em", 0.24)
            dlg_block_top_em = tp.get("dlg_block_top_em", 0.50)
            dlg_block_bottom_em = tp.get("dlg_block_bottom_em", 0.70)
        else:
            indent_em, para_gap_em, dialog_gap_em = 1.5, 0.6, 0.7
            h1_mb_em, h1_ls_em, dlg_side_pad_em = 1.2, -0.016, 0.0
            dlg_inblock_gap_em, dlg_block_top_em, dlg_block_bottom_em = 0.24, 0.50, 0.70

        css = f"""
/* === Readability+Typography by trim: {w_mm}mm (auto_tune={str(auto_tune).lower()}) === */
html, body {{
  margin: 0;
  padding: 0;
}}

body {{
  font-family: {font_stack};
  font-size: {eff_em}em;
  line-height: {self.opts.line_spacing};
  max-width: {maxw_em}em;
  margin: 0 auto;
  padding: {pad_em}em {pad_em}em;
  color: #1a1a1a;
  word-break: keep-all;
  overflow-wrap: break-word;
  -epub-hyphens: none;
  hyphens: none;
}}

h1 {{
  text-align: center;
  margin: 0 0 {h1_mb_em}em 0;
  font-weight: 800;
  letter-spacing: {h1_ls_em}em;
}}

p {{
  margin: 0 0 {para_gap_em}em 0;
  text-indent: {indent_em}em;
}}

p.dialog {{
  text-indent: 0;
  margin: 0 0 {dialog_gap_em}em 0;
  padding-left: {dlg_side_pad_em}em;
  padding-right: {dlg_side_pad_em}em;
}}

/* ✅ 대사 블록 자동 간격 */
p.dialog.dlg-start {{
  margin-top: {dlg_block_top_em}em;
}}
p.dialog.dlg-cont {{
  margin-top: 0;
}}
p.dialog.dlg-end {{
  margin-bottom: {dlg_block_bottom_em}em;
}}
p.dialog.dlg-cont:not(.dlg-end) {{
  margin-bottom: {dlg_inblock_gap_em}em;
}}
p.dialog.dlg-start.dlg-end {{
  margin-top: {dlg_block_top_em}em;
  margin-bottom: {dlg_block_bottom_em}em;
}}

hr.scene {{
  border: 0;
  height: 1px;
  background: rgba(0,0,0,0.12);
  margin: 1.2em 0;
}}

@media (max-width: 520px) {{
  body {{
    max-width: 100%;
    padding: 1.0em 1.0em;
  }}
}}
"""
        return css.strip()

    # ==========================================================
    # EPUB helpers: used chars, subset, cover, repack, report
    # ==========================================================
    def _collect_used_chars(self, title: str, author: str, content: str) -> str:
        used = set(title + author + content)
        used.update(list(" \n\t\r"))
        used.update(list("“”‘’'\".,!?…—-()[]{}:;"))
        return "".join(sorted(used))

    def _subset_font_internal(self, font_path: str, used_chars: str) -> Tuple[Optional[bytes], str]:
        if TTFont is None or Subsetter is None or FtOptions is None:
            return (None, "")
        ext = Path(font_path).suffix.lower()
        try:
            if ext in (".ttc", ".otc"):
                font = TTFont(font_path, fontNumber=0)
                out_ext = ".ttf"
            else:
                font = TTFont(font_path)
                out_ext = ext if ext in (".ttf", ".otf") else ".ttf"
            options = FtOptions()
            options.desubroutinize = True
            subsetter = Subsetter(options=options)
            subsetter.populate(text=used_chars)
            subsetter.subset(font)
            import io
            buf = io.BytesIO()
            font.save(buf)
            return (buf.getvalue(), out_ext)
        except Exception:
            return (None, "")

    def _subset_font_external(self, font_path: str, used_chars: str, cmd_template: str) -> Tuple[Optional[bytes], str]:
        if not cmd_template.strip():
            return (None, "")

        tmp_dir = self.temp_dir / "subset"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        textfile = tmp_dir / "used_chars.txt"
        textfile.write_text(used_chars, encoding="utf-8")

        in_path = Path(font_path)
        out_path = tmp_dir / (in_path.stem + ".subset.ttf")

        cmd = (cmd_template
               .replace("{in}", str(in_path))
               .replace("{out}", str(out_path))
               .replace("{textfile}", str(textfile)))

        try:
            proc = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if proc.returncode != 0:
                raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "외부 서브셋 실패")
            if out_path.exists() and out_path.stat().st_size > 0:
                return (out_path.read_bytes(), ".ttf")
        except Exception:
            return (None, "")

        return (None, "")

    def _make_cover_bytes(self, title: str, author: str) -> Tuple[Optional[str], Optional[bytes]]:
        if Image is None:
            return (None, None)

        fmt = self.opts.epub_opt.cover_format.upper()
        max_px = int(self.opts.epub_opt.cover_max_px)

        w, h = 1200, 1600
        img = Image.new("RGB", (w, h), (245, 245, 245))
        draw = ImageDraw.Draw(img)

        font_path = self.epub_font_path()
        try:
            if font_path and os.path.exists(font_path):
                title_font = ImageFont.truetype(font_path, 72)
                author_font = ImageFont.truetype(font_path, 40)
            else:
                title_font = ImageFont.load_default()
                author_font = ImageFont.load_default()
        except Exception:
            title_font = ImageFont.load_default()
            author_font = ImageFont.load_default()

        draw.rectangle([70, 70, w - 70, h - 70], outline=(200, 200, 200), width=6)
        draw.text((120, 380), title or "제목 없음", fill=(40, 40, 40), font=title_font)
        draw.text((120, 520), f"by {author or '작자 미상'}", fill=(90, 90, 90), font=author_font)

        if max_px and max(w, h) > max_px:
            scale = max_px / float(max(w, h))
            nw, nh = int(w * scale), int(h * scale)
            img = img.resize((nw, nh), resample=Image.Resampling.LANCZOS)

        import io
        buf = io.BytesIO()
        if fmt == "PNG":
            img.save(buf, format="PNG", optimize=True)
            return ("cover.png", buf.getvalue())
        else:
            q = int(self.opts.epub_opt.cover_jpeg_quality)
            img.save(buf, format="JPEG", quality=q, optimize=True, progressive=True)
            return ("cover.jpg", buf.getvalue())

    def _run_external_epub_tool(self, in_epub: str, cmd_template: str) -> str:
        if not cmd_template.strip():
            return in_epub
        out_epub = str(Path(in_epub).with_suffix(".opt.epub"))
        cmd = cmd_template.replace("{in}", in_epub).replace("{out}", out_epub)
        try:
            proc = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if proc.returncode != 0:
                raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "외부 EPUB 도구 실패")
            if os.path.exists(out_epub) and os.path.getsize(out_epub) > 0:
                return out_epub
        except Exception as e:
            QMessageBox.warning(self, "외부 EPUB 도구 실패", str(e))
        return in_epub

    def _repack_epub_zip(self, epub_path: str, compresslevel: int = 9):
        src = Path(epub_path)
        tmp = src.with_suffix(".repack.epub")

        with zipfile.ZipFile(src, "r") as zin:
            names = zin.namelist()
            with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=compresslevel) as zout:
                # mimetype first, stored
                mimetype_data = b"application/epub+zip"
                if "mimetype" in names:
                    try:
                        mimetype_data = zin.read("mimetype").strip() or mimetype_data
                    except Exception:
                        pass
                info = zipfile.ZipInfo("mimetype")
                info.compress_type = zipfile.ZIP_STORED
                zout.writestr(info, mimetype_data)

                for name in names:
                    if name == "mimetype":
                        continue
                    data = zin.read(name)
                    zout.writestr(name, data)

        tmp.replace(src)

    def _epub_size_report(self, epub_path: str, top_n: int = 12) -> str:
        p = Path(epub_path)
        if not p.exists():
            return "리포트 생성 실패: 파일이 없습니다."

        cat = {"fonts": 0, "images": 0, "html": 0, "css": 0, "meta": 0, "other": 0}
        entries: List[Tuple[str, int]] = []

        with zipfile.ZipFile(p, "r") as z:
            for info in z.infolist():
                name = info.filename
                size = int(info.file_size)
                entries.append((name, size))

                low = name.lower()
                if low.endswith((".ttf", ".otf", ".ttc", ".otc", ".woff", ".woff2")):
                    cat["fonts"] += size
                elif low.endswith((".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg")):
                    cat["images"] += size
                elif low.endswith((".xhtml", ".html", ".htm")):
                    cat["html"] += size
                elif low.endswith(".css"):
                    cat["css"] += size
                elif low.startswith("meta-inf") or low.endswith((".opf", ".ncx")) or low == "mimetype":
                    cat["meta"] += size
                else:
                    cat["other"] += size

        total = sum(cat.values())
        entries.sort(key=lambda x: x[1], reverse=True)
        top = entries[:top_n]

        def kb(n: int) -> str:
            return f"{n/1024:.1f} KB"

        lines = [
            f"[EPUB 용량 리포트] {p.name}",
            f"총합: {kb(total)}",
            "",
            "카테고리별:",
            f" - fonts: {kb(cat['fonts'])}",
            f" - images: {kb(cat['images'])}",
            f" - html: {kb(cat['html'])}",
            f" - css: {kb(cat['css'])}",
            f" - meta: {kb(cat['meta'])}",
            f" - other: {kb(cat['other'])}",
            "",
            f"상위 {len(top)}개 파일:",
        ]
        for name, size in top:
            lines.append(f" - {name} ({kb(size)})")
        return "\n".join(lines)

    # ==========================================================
    # EPUB save (판형/타이포 자동 + 대사 블록 자동)
    # ==========================================================
    def save_as_epub(self, path: str, content: str):
        if not epub:
            raise Exception("EbookLib 라이브러리가 필요합니다. (pip install ebooklib)")

        title = (self.title_edit.text().strip() or "제목 없음")
        author = (self.author_edit.text().strip() or "작자 미상")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        # ✅ 부작용 방지(저장 중 opt 변형 금지)
        opt = replace(self.opts.epub_opt)
        prof = self.opts.epub_profile

        # 프리셋 반영(저장 시 안전하게)
        preset = READER_PRESETS.get(opt.reader_preset, {})
        embed_font = bool(preset.get("embed_font", prof.embed_font)) if opt.optimize_mode else bool(prof.embed_font)

        if opt.optimize_mode and preset:
            opt.cover_max_px = int(preset.get("cover_max_px", opt.cover_max_px))
            opt.cover_format = str(preset.get("cover_format", opt.cover_format))
            opt.cover_jpeg_quality = int(preset.get("cover_jpeg_quality", opt.cover_jpeg_quality))
            opt.repack_zip = bool(preset.get("repack_zip", opt.repack_zip))
            opt.zip_compresslevel = int(preset.get("zip_level", opt.zip_compresslevel))

        book = epub.EpubBook()
        book.set_title(title)
        book.add_author(author)
        book.set_language("ko")

        base_em = prof.base_font_em

        # CSS / font embed
        font_face_css = ""
        css_item = None

        if embed_font:
            font_path = self.epub_font_path()
            if font_path and os.path.exists(font_path):
                used_chars = self._collect_used_chars(title, author, content)
                font_bytes = None
                out_ext = Path(font_path).suffix.lower()

                if opt.subset_font:
                    if opt.prefer_internal_subsetter:
                        font_bytes, out_ext = self._subset_font_internal(font_path, used_chars)

                    if font_bytes is None and opt.external_font_subset_enabled and opt.external_font_subset_cmd.strip():
                        font_bytes, out_ext = self._subset_font_external(font_path, used_chars, opt.external_font_subset_cmd)

                    if font_bytes is None:
                        font_bytes = Path(font_path).read_bytes()
                        out_ext = Path(font_path).suffix.lower()
                else:
                    font_bytes = Path(font_path).read_bytes()
                    out_ext = Path(font_path).suffix.lower()

                media = "font/ttf" if out_ext in (".ttf", ".ttc") else "font/otf"
                if out_ext == ".ttf":
                    media = "font/ttf"

                font_item_name = f"fonts/novel{out_ext if out_ext else '.ttf'}"
                font_item = epub.EpubItem(
                    uid="novel_font",
                    file_name=font_item_name,
                    media_type=media,
                    content=font_bytes
                )
                book.add_item(font_item)

                font_face_css = (
                    "@font-face{"
                    "font-family:'NovelFont';"
                    f"src:url('../{font_item_name}');"
                    "}"
                )

        w_mm, h_mm = self.paper_format_mm()

        if embed_font:
            font_stack = "'NovelFont', serif"
        else:
            font_stack = f"'{escape(prof.family)}','Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif"

        readability_css = self._epub_readability_css_by_trim(
            w_mm=w_mm,
            base_em=base_em,
            auto_tune=bool(opt.auto_typo_tune),
            font_stack=font_stack
        )

        style = (
            font_face_css
            + readability_css
            + """
.titlebox{margin-top:18vh; text-align:center;}
.meta{color:#666; font-size:0.95em; margin-top:10px;}
"""
        )

        css_item = epub.EpubItem(uid="style", file_name="style/main.css",
                                 media_type="text/css", content=style.encode("utf-8"))
        book.add_item(css_item)

        # Title page + cover
        titlepage = None
        if self.opts.include_title_page:
            cover_name, cover_bytes = self._make_cover_bytes(title, author)
            if cover_name and cover_bytes:
                book.set_cover(cover_name, cover_bytes, create_page=True)

            titlepage = epub.EpubHtml(title="Title Page", file_name="titlepage.xhtml", lang="ko")
            titlepage.add_item(css_item)
            titlepage.content = (
                f"<div class='titlebox'>"
                f"<h1>{escape(title)}</h1>"
                f"<div class='meta'>저자: {escape(author)}</div>"
                f"<div class='meta'>생성일: {escape(now_str)}</div>"
                f"</div>"
            )
            book.add_item(titlepage)

        # Chapters
        chapters_data = split_into_chapters(content)
        epub_chapters = []

        for ch_info in chapters_data:
            tokens = smart_make_paragraphs(ch_info["lines"])
            items = []
            for t in tokens:
                if not t.strip():
                    items.append({"kind": "scene"})
                    continue
                s = t
                is_dlg = s.lstrip().startswith(DIALOGUE_PREFIXES)
                items.append({"kind": "p", "text": s, "is_dialogue": is_dlg})

            if not any(it["kind"] == "p" for it in items):
                continue

            ch = epub.EpubHtml(
                title=ch_info["title"],
                file_name=f"chap_{len(epub_chapters):03}.xhtml",
                lang="ko"
            )
            ch.add_item(css_item)

            html_parts = [f"<h1>{escape(ch_info['title'])}</h1>"]

            for i, it in enumerate(items):
                if it["kind"] == "scene":
                    html_parts.append("<hr class='scene' />")
                    continue

                text = it["text"]
                is_dlg = it["is_dialogue"]

                if not is_dlg:
                    html_parts.append(f"<p>{escape(text)}</p>")
                    continue

                prev_is_dlg = (i > 0 and items[i - 1]["kind"] == "p" and items[i - 1].get("is_dialogue") is True)
                next_is_dlg = (i + 1 < len(items) and items[i + 1]["kind"] == "p" and items[i + 1].get("is_dialogue") is True)

                classes = ["dialog"]
                classes.append("dlg-cont" if prev_is_dlg else "dlg-start")
                if not next_is_dlg:
                    classes.append("dlg-end")

                cls_attr = " ".join(classes)
                html_parts.append(f"<p class='{cls_attr}'>{escape(text)}</p>")

            ch.content = "".join(html_parts)

            book.add_item(ch)
            epub_chapters.append(ch)

        if not epub_chapters:
            ch = epub.EpubHtml(title="본문", file_name="chap_000.xhtml", lang="ko")
            ch.add_item(css_item)
            ch.content = "<h1>본문</h1><p>내용이 없습니다.</p>"
            book.add_item(ch)
            epub_chapters.append(ch)

        spine = ["nav"]
        if titlepage:
            spine.append(titlepage)
        spine.extend(epub_chapters)
        book.spine = spine

        if titlepage:
            book.toc = (
                epub.Link(titlepage.file_name, "Title Page", "titlepage"),
                epub.Section("본문", epub_chapters),
            )
        else:
            book.toc = (epub.Section("본문", epub_chapters),)

        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())

        # 1) 기본 저장
        epub.write_epub(path, book, {})

        # 2) 외부 EPUB 최적화(선택)
        if opt.optimize_mode and opt.external_epub_tool_enabled and opt.external_epub_tool_cmd.strip():
            out = self._run_external_epub_tool(path, opt.external_epub_tool_cmd)
            if out != path and os.path.exists(out):
                shutil.copy2(out, path)
                try:
                    os.remove(out)
                except Exception:
                    pass

        # 3) ZIP 재압축(권장)
        if opt.optimize_mode and opt.repack_zip:
            self._repack_epub_zip(path, compresslevel=int(opt.zip_compresslevel))

        # 4) 용량 리포트
        report = self._epub_size_report(path)
        QMessageBox.information(self, "EPUB 용량 리포트", report)

    # ==========================================================
    # UI state save/restore
    # ==========================================================
    def restore_ui_state(self):
        geo = self.settings.value("ui/geometry", None)
        if geo is not None:
            try:
                self.restoreGeometry(geo)
            except Exception:
                pass

        spl = self.settings.value("ui/splitter", None)
        if spl is not None:
            try:
                self.splitter.restoreState(spl)
            except Exception:
                pass

    def save_ui_state(self):
        try:
            self.settings.setValue("ui/geometry", self.saveGeometry())
        except Exception:
            pass
        try:
            self.settings.setValue("ui/splitter", self.splitter.saveState())
        except Exception:
            pass
        self.settings.setValue("ui/last_url", self.url_edit.text().strip())
        self.settings.setValue("ui/last_format", self.format_combo.currentIndex())

    def closeEvent(self, event):
        self.save_ui_state()
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        event.accept()


# ==========================================================
# Run
# ==========================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())            