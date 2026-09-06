# -*- coding: utf-8 -*-
"""
EPUB Reader V2.14 (EPUB/TXT only)
- 썸네일 숨김/표시에도 페이지 중앙 틈이 생기지 않음
- 페이지 번호(캡션) Bold 제거
- 실행/세션복원 시 초기 용지 크기(설정값) 기준으로 먼저 페이징
- 폰트 설정값이 페이지 본문에 확실히 적용되도록(문서 defaultStyleSheet 강제)

Python 3.8+ / PySide6
"""

import os
import sys
import json
import time
import html as _html
import zipfile
import re
import traceback
from dataclasses import dataclass
from collections import OrderedDict, deque
from typing import Optional, List, Tuple, Dict, Any

from xml.etree import ElementTree as ET

from PySide6.QtCore import (
    Qt, QSize, QSizeF, QRect, QRectF, QPoint, QEvent, QTimer,
    QPropertyAnimation, QEasingCurve, Signal, QVariantAnimation
)
from PySide6.QtGui import (
    QAction, QActionGroup, QFont, QFontDatabase, QColor, QCursor,
    QKeySequence, QPixmap, QPainter, QIcon, QTextDocument, QTextOption,
    QTextCursor, QTextCharFormat, QRadialGradient, QPainterPath, QLinearGradient, QPen
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QLabel, QSplitter, QFrame, QComboBox,
    QSpinBox, QDoubleSpinBox, QMessageBox, QFileDialog, QInputDialog,
    QFontDialog, QToolButton, QStyle, QStackedWidget, QMenuBar, QMenu,
    QTextBrowser, QTextEdit, QProgressBar, QPushButton
)

SETTINGS_FILE = "epub_reader_v2_13_settings.json"
SUPPORTED_OPEN_EXT = {".epub", ".txt"}

THEMES = {
    "화이트": "#ffffff",
    "블랙": "#111111",
    "종이":  "#fdf6e3",
}

CANVAS_BG = "#8a8f98"

PAPER_PRESETS = {
    "전체보기(창맞춤)": None,
    "A5 (소설)": (148, 210),
    "B6 (문고)": (128, 182),
    "Pocket (포켓)": (105, 148),
    "Kindle 6\"": (108, 144),
}

BAD_FONTS = {"Fixedsys", "System", "Terminal", "Small Fonts", "MS Sans Serif", "8514oem"}
DEFAULT_TEXT_PAPER = QSize(760, 980)

TEXT_THUMB_CACHE_MAX = 250
THUMB_ITEM_SIZE = QSize(180, 240)
THUMB_BATCH = 10
THUMB_PREFETCH_PAD = 30
THUMB_FAST_CENTER_PAD = 12

WHEEL_ANGLE_THRESHOLD = 120
WHEEL_PIXEL_THRESHOLD = 120

DEFAULT_USER_HIGHLIGHT_COLOR = "#ffe082"
DEFAULT_SEARCH_HIGHLIGHT_COLOR = "#b3e5fc"

PAGE_TURN_MS = 320
THUMB_SCROLL_ANIM_MS = 220


def clamp(v: int, a: int, b: int) -> int:
    return max(a, min(b, v))


def safe_ui_font(preferred: Tuple[str, ...] = ("맑은 고딕", "Segoe UI", "Noto Sans CJK KR", "Arial")) -> QFont:
    fams = set(QFontDatabase.families())
    for fam in preferred:
        if fam in fams and fam not in BAD_FONTS:
            f = QFont(fam)
            f.setStyleHint(QFont.SansSerif)
            return f
    f = QApplication.font()
    f.setStyleHint(QFont.SansSerif)
    return f


def is_korean_text(text: str) -> bool:
    if not text:
        return True
    hangul_count = len(re.findall(r"[\uAC00-\uD7A3]", text))
    english_count = len(re.findall(r"[a-zA-Z]", text))
    if hangul_count == 0 and english_count > 0:
        return False
    total = hangul_count + english_count
    if total > 0 and (hangul_count / total) < 0.05:
        return False
    return True


def detect_best_encoding(data: bytes) -> str:
    """
    바이너리 데이터의 인코딩을 최적으로 자동 감지한다.
    BOM 감지 -> Strict 디코딩 시도 (latin-1 제외) -> 점수식 한글 감지 순.
    """
    # 1. BOM 감지
    if data[:3] == b"\xef\xbb\xbf":
        return "utf-8-bom"
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "utf-16"
    if data[:4] in (b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff"):
        return "utf-32"

    # 2. Strict 디코딩 시도 (손상되지 않은 완벽한 파일인 경우 우선순위대로 즉각 로드. 단, latin-1은 무조건 성공하므로 제외)
    strict_encodings = ["utf-8", "cp949", "euc-kr", "utf-16"]
    for enc in strict_encodings:
        try:
            data.decode(enc)
            return enc
        except (UnicodeDecodeError, LookupError):
            continue

    # 3. 완벽한 디코딩이 불가능한 경우 (파일 내부 일부 손상 또는 복합 바이트)
    # 한글 빈도수 및 대체 문자(errors='replace' 시 \ufffd) 개수를 점수화하여 최적의 인코딩 선택.
    # 영어 및 서양어 처리를 위해 latin-1도 후보군에 포함합니다.
    score_encodings = ["utf-8", "cp949", "euc-kr", "utf-16", "latin-1"]
    hangul_re = re.compile(r"[\uac00-\ud7a3\u1100-\u11ff\u3130-\u318f]")
    max_score = -99999999
    best_enc = "utf-8"

    for enc in score_encodings:
        try:
            decoded_text = data.decode(enc, errors="replace")
            replace_count = decoded_text.count("\ufffd")
            hangul_count = len(hangul_re.findall(decoded_text))

            # 한글 1글자당 100점 가중치, 대체 문자 1글자마다 1점 감점
            score = (hangul_count * 100) - replace_count

            if score > max_score:
                max_score = score
                best_enc = enc
        except LookupError:
            continue

    return best_enc


@dataclass
class TextLayoutSettings:
    font_family: str = "맑은 고딕"
    font_size: int = 13
    bold: bool = False
    line_spacing: float = 1.25
    para_spacing: int = 6
    indent: float = 1.5
    margin_lr: int = 32
    margin_tb: int = 36


class NoScrollTextBrowser(QTextBrowser):
    """페이지 내부 스크롤 완전 차단"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.setFocusPolicy(Qt.ClickFocus)
        self.setOpenExternalLinks(False)
        self.setFrameShape(QFrame.NoFrame)
        self.viewport().setAutoFillBackground(True)

    def wheelEvent(self, e):
        e.ignore()

    def scrollContentsBy(self, dx: int, dy: int):
        return

    def ensureCursorVisible(self):
        return

    def keyPressEvent(self, e):
        k = e.key()
        if k in (
            Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right,
            Qt.Key_PageUp, Qt.Key_PageDown, Qt.Key_Home, Qt.Key_End,
            Qt.Key_Space
        ):
            e.ignore()
            return
        super().keyPressEvent(e)


class ThumbListWidget(QListWidget):
    """썸네일에서 휠로 선택 이동: 1장=1칸, 2장=2칸"""
    visibleRangeChanged = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._accum_angle = 0
        self._accum_pixel = 0
        self.setUniformItemSizes(True)

        sb = self.verticalScrollBar()
        sb.valueChanged.connect(lambda _v: self.visibleRangeChanged.emit())

    def wheelEvent(self, e):
        if self.count() <= 0:
            e.accept()
            return

        dy_pixel = 0
        dy_angle = 0
        try:
            if hasattr(e, "pixelDelta") and not e.pixelDelta().isNull():
                dy_pixel = int(e.pixelDelta().y())
            dy_angle = int(e.angleDelta().y())
        except Exception:
            pass

        moved = False
        direction = 0

        if dy_pixel != 0:
            self._accum_pixel += dy_pixel
            if abs(self._accum_pixel) >= WHEEL_PIXEL_THRESHOLD:
                direction = -1 if self._accum_pixel > 0 else 1
                moved = True
                self._accum_pixel = 0
                self._accum_angle = 0
        else:
            self._accum_angle += dy_angle
            if abs(self._accum_angle) >= WHEEL_ANGLE_THRESHOLD:
                direction = -1 if self._accum_angle > 0 else 1
                moved = True
                self._accum_angle = 0
                self._accum_pixel = 0

        if not moved:
            e.accept()
            return

        main = self.window()
        unit = getattr(main, "view_mode", 1)
        if unit not in (1, 2):
            unit = 1
        step = direction * unit

        cur = self.currentRow()
        if cur < 0:
            cur = 0

        new_row = max(0, min(self.count() - 1, cur + step))
        if unit == 2:
            new_row = new_row - (new_row % 2)

        if new_row != cur:
            self.setCurrentRow(new_row)
            it = self.item(new_row)
            if it:
                self.scrollToItem(it, QListWidget.PositionAtCenter)

        self.visibleRangeChanged.emit()
        e.accept()


class TextPager:
    def __init__(self, layout: Optional[TextLayoutSettings] = None, paper_px: Optional[QSize] = None):
        self.text = ""
        self.pages: List[str] = []
        self.page_starts: List[int] = []
        self.page_continuations: List[bool] = []  # 각 페이지가 이전 단락 이어짐 여부
        self.layout = layout or TextLayoutSettings()
        self.paper_px = QSize(paper_px or DEFAULT_TEXT_PAPER)
        self._last_sig = None
        self.fg = "#111111"
        self.bg = "#ffffff"
        self.progress_callback = None
        self.is_korean = True

    def set_colors(self, fg: str, bg: str):
        self.fg = fg
        self.bg = bg

    def signature(self) -> Tuple[Any, ...]:
        return (
            int(self.paper_px.width()),
            int(self.paper_px.height()),
            self.layout.font_family,
            int(self.layout.font_size),
            bool(self.layout.bold),
            float(self.layout.line_spacing),
            int(self.layout.para_spacing),
            int(self.layout.margin_lr),
            int(self.layout.margin_tb),
            len(self.text),
            self.fg,
            self.bg,
            bool(getattr(self, "is_korean", True)),
        )

    # ✅ CSS를 문서 defaultStyleSheet로도 강제 적용하기 위해 분리
    def build_css(self, layout: Optional[TextLayoutSettings] = None) -> str:
        layout = layout or self.layout
        weight = "700" if layout.bold else "400"
        
        is_kor = getattr(self, "is_korean", True)
        if not is_kor:
            line_h = max(1.0, float(layout.line_spacing) * 0.90)
            para_spacing = max(0, int(layout.para_spacing * 0.2))
            w_break = "normal"
            l_break = "normal"
            t_justify = "auto"
            align_last = "left"
            o_wrap = "break-word"
        else:
            line_h = max(1.0, float(layout.line_spacing))
            para_spacing = layout.para_spacing
            w_break = "keep-all"
            l_break = "anywhere"
            t_justify = "inter-character"
            align_last = "justify"
            o_wrap = "anywhere"

        return f"""
        html, body {{ margin:0; padding:0; }}
        body, p {{
            font-family: '{layout.font_family}';
            font-size: {layout.font_size}pt;
            font-weight: {weight};
            line-height: {line_h};
            color: {self.fg};
        }}
        body {{
            background: {self.bg};
            text-align: justify;
            text-align-last: {align_last};
            text-justify: {t_justify};
            line-break: {l_break};
            word-break: {w_break};
            overflow-wrap: {o_wrap};
            hyphens: none;
            letter-spacing: 0px;
            word-spacing: 0px;
        }}
        p {{
            margin: 0 0 {para_spacing}px 0;
            text-indent: {layout.indent}em;
            text-align: justify;
            text-align-last: {align_last};
        }}
        p:last-child {{
            margin-bottom: 0;
        }}
        """

    def _insert_korean_breaks(self, text: str) -> str:
        # 한글 사이에만 zero-width space를 삽입하여 불필요한 줄바꿈을 줄입니다.
        return re.sub(
            r'(?<=[\uAC00-\uD7A3\u1100-\u11FF\u3130-\u318F])(?=[\uAC00-\uD7A3\u1100-\u11FF\u3130-\u318F])',
            "\u200b",
            text,
        )

    # 페이징 진행률 콜백을 저장할 수 있도록 속성 추가
    progress_callback: Optional[callable] = None

    def _prepare_text_for_paging(self, raw: str) -> str:
        """
        페이징 전용 텍스트 전처리:
        - Unicode 정규화 및 제로-폭 문자 제거
        - 문단 구분(빈 줄 및 들여쓰기 시작)을 안전하게 보존
        - 문장 중간의 불필요한 개행 제거 및 공백 축약
        - 문장 부호 뒤에 페이징용 줄바꿈 삽입
        """
        import unicodedata

        if not raw:
            return raw

        # 1. 유니코드 정규화 및 제로-폭 문자 제거
        s = unicodedata.normalize("NFKC", raw)
        s = re.sub(r"[\u200B\u200C\u200D\uFEFF\u2060]", "", s)

        # 줄바꿈 통일
        s = s.replace("\r\n", "\n")

        # 2. 문단 구분(연속 개행, 개행+들여쓰기 공백)을 임시 마커로 치환하여 안전 보존
        # 2-1. 연속 개행 (\n\n 등)
        s = re.sub(r"\n{2,}", "【PARA】", s)
        # 2-2. 개행 + 들여쓰기 공백
        s = re.sub(r"\n[ \t]+", "【PARA】", s)

        is_kor = getattr(self, "is_korean", True)
        if is_kor:
            # 한글 문서: 대화 시작/종료 기호 주변 개행을 문단 구분으로 정확히 보존하여 문단 간격을 띄움
            dialogue_chars = r"\"'“‘「『\-\(\[—"
            s = re.sub(rf"\n(?=[{dialogue_chars}])", "【PARA】", s)
            end_chars = r"\"'”’」』"
            s = re.sub(rf"(?<=[{end_chars}])\n", "【PARA】", s)

            # 한글 문장 중간 개행을 공백으로 합침
            hangul = r"[\uAC00-\uD7A3\u1100-\u11FF\u3130-\u318F]"
            s = re.sub(rf"(?<={hangul})\n(?={hangul})", " ", s)

        # 3. 마커가 아닌 나머지 단일 개행(\n)은 문장 중간의 단순 개행이므로 공백으로 합침
        s = s.replace("\n", " ")

        # 4. 여러 공백/탭을 하나로 축약
        s = re.sub(r"[ \t\f\v]+", " ", s)

        # 5. 임시 마커를 실제 문단 구분자인 \n\n으로 복원
        s = s.replace("【PARA】", "\n\n")

        # 연속된 개행 정리
        s = re.sub(r"\n{3,}", "\n\n", s)

        return s.strip()

    def build_html(self, text_chunk: str, is_continuation: bool = False, is_continued_to_next: bool = False) -> str:
        """
        is_continuation=True 이면 첫 번째 단락에 들여쓰기를 적용하지 않는다.
        is_continued_to_next=True 이면 마지막 단락의 마지막 줄까지 양쪽 정렬(justify)을 적용한다.
        """
        layout = self.layout
        css = self.build_css(layout)

        raw = (text_chunk or "").replace("\r\n", "\n")
        raw = self._insert_korean_breaks(raw)
        raw = re.sub(r"\n{2,}", "\n\n", raw)
        raw = _html.escape(raw, quote=False)

        paras = raw.split("\n\n")
        indent_spaces = max(0, int(round(layout.indent * 4)))
        indent_prefix = "&nbsp;" * indent_spaces
        parts = []
        is_kor = getattr(self, "is_korean", True)
        align_last = "left" if not is_kor else "justify"
        for idx, p in enumerate(paras):
            p_html = p

            # 이어지는 단락의 첫 번째 항목은 들여쓰기 없음
            if idx == 0 and is_continuation:
                parts.append(
                    f"<p style='text-indent:0; text-align:justify; text-align-last:{align_last};'>{p_html}</p>"
                )
            else:
                parts.append(
                    f"<p style='text-indent:{layout.indent}em; text-align:justify; text-align-last:{align_last};'>{indent_prefix}{p_html}</p>"
                )
        body = "".join(parts)
        return f"<html><head><style>{css}</style></head><body>{body}</body></html>"

    def load_txt(self, path: str, encoding: Optional[str] = None):
        """
        TXT 파일 로드. encoding을 지정하면 해당 인코딩으로 강제 디코딩.
        미지정 시 최적의 인코딩 자동 감지.
        """
        data = open(path, "rb").read()
        self.last_encoding = "unknown"

        if encoding:
            # 수동 지정 인코딩
            self.text = data.decode(encoding, errors="replace")
            self.last_encoding = encoding
        else:
            detected_enc = detect_best_encoding(data)
            self.last_encoding = detected_enc
            if detected_enc == "utf-8-bom":
                self.text = data[3:].decode("utf-8", errors="replace")
            else:
                self.text = data.decode(detected_enc, errors="replace")

        self.is_korean = is_korean_text(self.text)
        self._paginate()


    def load_text(self, text: str):
        self.text = text or ""
        self.is_korean = is_korean_text(self.text)
        self._paginate()

    def set_layout(self, layout: TextLayoutSettings):
        self.layout = layout
        self._paginate()

    def set_paper_px(self, size: QSize):
        size = QSize(max(220, size.width()), max(220, size.height()))
        if size == self.paper_px:
            return
        self.paper_px = QSize(size)
        self._paginate()

    def page_count(self) -> int:
        return len(self.pages)

    def _build_doc(self, text_chunk: str, paper_px: Optional[QSize] = None, layout: Optional[TextLayoutSettings] = None, is_continuation: bool = False, is_continued_to_next: bool = False) -> QTextDocument:
        layout = layout or self.layout
        doc = QTextDocument()
        doc.setDocumentMargin(0)

        f = QFont(layout.font_family, layout.font_size)
        f.setBold(layout.bold)
        f.setStyleHint(QFont.SansSerif)
        doc.setDefaultFont(f)
        doc.setUseDesignMetrics(True)

        # ✅ 문서 기본 스타일시트 강제(렌더링 일관성)
        doc.setDefaultStyleSheet(self.build_css(layout))

        opt = QTextOption()
        opt.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        opt.setAlignment(Qt.AlignJustify | Qt.AlignTop)
        doc.setDefaultTextOption(opt)

        doc.setHtml(self.build_html(text_chunk, is_continuation=is_continuation, is_continued_to_next=is_continued_to_next))

        paper = QSize(paper_px or self.paper_px)
        w = paper.width() - layout.margin_lr * 2
        h = paper.height() - layout.margin_tb * 2 - 20

        SAFE_PAD = 16
        page_w = max(100, w)
        page_h = max(100, h - SAFE_PAD)

        doc.setTextWidth(page_w)
        doc.setPageSize(QSizeF(page_w, page_h))
        return doc

    def _doc_real_height(self, doc: QTextDocument) -> float:
        try:
            return float(doc.documentLayout().documentSize().height())
        except Exception:
            return float(doc.size().height())

    def _paginate(self, force: bool = False):
        sig = self.signature()
        if (not force) and sig == self._last_sig:
            return
        self._last_sig = sig

        self.pages.clear()
        self.page_starts.clear()
        self.page_continuations.clear()

        raw = (self.text or "").replace("\r\n", "\n")
        # 페이징 전에 텍스트 정리(제로-폭 제거, 분할된 한글 복원, 문장부호 기준 줄바꿈)
        raw = self._prepare_text_for_paging(raw)
        if not raw.strip():
            return
        total_len = len(raw)
        last_report = -1

        approx = max(900, int((self.paper_px.width() * self.paper_px.height()) / 900))
        i, n = 0, len(raw)
        MIN_CHUNK = 250
        # 80자 창에서 단어 경계 탐색
        WIN = 80
        next_is_continuation = False  # 다음 페이지가 단락 이어짐인지

        while i < n:
            chunk = raw[i:i + approx]
            if not chunk:
                break

            is_cont = next_is_continuation
            next_is_continuation = False

            doc = self._build_doc(chunk, is_continuation=is_cont)
            max_h = float(doc.pageSize().height())
            real_h = self._doc_real_height(doc)

            if real_h <= max_h:
                best = len(chunk)
                max_extend = min(n - i, max(best * 3, approx * 2, 2000))
                lo = best + 1
                hi = max_extend
                while lo <= hi:
                    mid = (lo + hi) // 2
                    td = self._build_doc(raw[i:i + mid], is_continuation=is_cont)
                    if self._doc_real_height(td) <= max_h:
                        best = mid
                        lo = mid + 1
                    else:
                        hi = mid - 1

                best = max(best, MIN_CHUNK)
                cut_text = raw[i:i + best]
                at_para = cut_text.rstrip().endswith("\n") or raw[i + best:i + best + 2] == "\n\n"

                chunk = raw[i:i + best]
                if not at_para:
                    next_is_continuation = True

                # 페이지 시작에 불필요한 빈 줄이 나오지 않도록 선행 개행 제거 후 저장
                stored_chunk = chunk.lstrip('\n')
                self.page_starts.append(i)
                self.pages.append(stored_chunk)
                self.page_continuations.append(is_cont)
                i += len(chunk)
                # 페이징 진행률 보고
                try:
                    if self.progress_callback and total_len:
                        percent = int((i / total_len) * 100)
                        if percent != last_report:
                            self.progress_callback(percent, '페이징 중')
                            last_report = percent
                            QApplication.processEvents()
                except Exception:
                    pass
                continue

            lo, hi = MIN_CHUNK, len(chunk)
            best = MIN_CHUNK
            while lo <= hi:
                mid = (lo + hi) // 2
                td = self._build_doc(chunk[:mid], is_continuation=is_cont)
                if self._doc_real_height(td) <= max_h:
                    best = mid
                    lo = mid + 1
                else:
                    hi = mid - 1

            best = max(1, best)
            cut_text = chunk[:best]
            at_para = cut_text.rstrip().endswith("\n") or raw[i + best:i + best + 2] == "\n\n"

            if not at_para:
                next_is_continuation = True

            segment = chunk[:best]
            # 페이지 시작에 불필요한 빈 줄이 나오지 않도록 선행 개행 제거 후 저장
            stored_segment = segment.lstrip('\n')
            self.page_starts.append(i)
            self.pages.append(stored_segment)
            self.page_continuations.append(is_cont)
            i += best
            # 페이징 진행률 보고
            try:
                if self.progress_callback and total_len:
                    percent = int((i / total_len) * 100)
                    if percent != last_report:
                        self.progress_callback(percent, '페이징 중')
                        last_report = percent
                        QApplication.processEvents()
            except Exception:
                pass

    def render_thumb_icon(self, page_idx: int, bg: QColor) -> QIcon:
        if page_idx < 0 or page_idx >= len(self.pages):
            return QIcon()

        # ✅ 썸네일 전용 레이아웃: 더 작은 글자/여백으로 더 많은 텍스트를 수용
        THUMB_SCALE = 0.72  # 0.70~0.80 추천
        thumb_layout = TextLayoutSettings(
            font_family=self.layout.font_family,
            font_size=max(7, int(self.layout.font_size * 0.80)),
            bold=self.layout.bold,
            line_spacing=max(1.0, self.layout.line_spacing * 0.90),
            para_spacing=max(4, int(self.layout.para_spacing * 0.5)),
            margin_lr=max(6, int(self.layout.margin_lr * 0.25)),
            margin_tb=max(6, int(self.layout.margin_tb * 0.25)),
        )

        thumb_paper = QSize(
            int(THUMB_ITEM_SIZE.width() / THUMB_SCALE),
            int(THUMB_ITEM_SIZE.height() / THUMB_SCALE),
        )
        doc = self._build_doc(self.pages[page_idx], paper_px=thumb_paper, layout=thumb_layout)

        pm = QPixmap(THUMB_ITEM_SIZE)
        pm.fill(bg)

        painter = QPainter(pm)
        painter.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        painter.scale(THUMB_SCALE, THUMB_SCALE)
        painter.translate(thumb_layout.margin_lr, thumb_layout.margin_tb)
        doc.drawContents(painter, QRectF(0, 0, doc.pageSize().width(), doc.pageSize().height()))
        painter.end()

        return QIcon(pm)


def _strip_html_tags(s: str) -> str:
    """
    EPUB XHTML/HTML에서 텍스트만 추출:
    - <br>, </p> 등은 개행으로 변환
    - 엔티티 해제는 태그 제거 이후에 수행
    """
    if not s:
        return ""

    s = re.sub(r"(?is)<script.*?>.*?</script>", " ", s)
    s = re.sub(r"(?is)<style.*?>.*?</style>", " ", s)
    s = re.sub(r"(?is)<br\s*/?>", "\n", s)
    s = re.sub(r"(?is)</p\s*>", "\n\n", s)
    s = re.sub(r"(?is)</div\s*>", "\n", s)
    s = re.sub(r"(?is)</h[1-6]\s*>", "\n\n", s)
    s = re.sub(r"(?is)<.*?>", " ", s)

    s = _html.unescape(s)

    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def extract_text_from_epub(epub_path: str, progress_callback=None) -> str:
    """
    표준 container.xml → OPF → manifest/spine 순서로 XHTML/HTML을 찾아 텍스트 추출.
    """
    with zipfile.ZipFile(epub_path, "r") as z:
        try:
            container = z.read("META-INF/container.xml")
        except KeyError:
            return ""

        try:
            root = ET.fromstring(container)
        except Exception:
            return ""

        opf_path = None
        for el in root.iter():
            if el.tag.endswith("rootfile"):
                opf_path = el.attrib.get("full-path")
                break
        if not opf_path:
            return ""

        try:
            opf_data = z.read(opf_path)
        except KeyError:
            return ""

        try:
            opf_root = ET.fromstring(opf_data)
        except Exception:
            return ""

        manifest: Dict[str, str] = {}
        spine_ids: List[str] = []

        base_dir = os.path.dirname(opf_path).replace("\\", "/")
        if base_dir:
            base_dir += "/"

        for el in opf_root.iter():
            if el.tag.endswith("item"):
                _id = el.attrib.get("id")
                href = el.attrib.get("href")
                if _id and href:
                    manifest[_id] = (base_dir + href).replace("//", "/")
            if el.tag.endswith("itemref"):
                rid = el.attrib.get("idref")
                if rid:
                    spine_ids.append(rid)

        chunks: List[str] = []
        total = len(spine_ids)
        for idx, sid in enumerate(spine_ids):
            href = manifest.get(sid)
            if not href:
                continue
            low = href.lower()
            if not low.endswith((".xhtml", ".html", ".htm")):
                continue
            try:
                data = z.read(href)
            except KeyError:
                continue

            # 진행 콜백 호출
            try:
                if progress_callback and total:
                    percent = int(((idx + 1) / total) * 100)
                    progress_callback(percent, f"읽는중: {os.path.basename(href)}")
                    QApplication.processEvents()
            except Exception:
                pass

            detected_enc = detect_best_encoding(data)
            if detected_enc == "utf-8-bom":
                txt = data[3:].decode("utf-8", errors="replace")
            else:
                txt = data.decode(detected_enc, errors="replace")

            plain = _strip_html_tags(txt)
            if plain:
                chunks.append(plain)

        return "\n\n".join(chunks).strip()


class RightHintMenu(QMenu):
    def showEvent(self, e):
        super().showEvent(e)
        try:
            parent = self.parentWidget()
            if not parent:
                return
            screen = QApplication.screenAt(QCursor.pos()) if hasattr(QApplication, "screenAt") else None
            screen = screen or QApplication.primaryScreen()
            avail = screen.availableGeometry()
            if isinstance(parent, QMenu):
                pgeo = parent.geometry()
                desired_x = pgeo.right() + 2
                desired_y = self.y()
                if desired_x + self.width() <= avail.right():
                    self.move(desired_x, desired_y)
        except Exception:
            pass


class SmartMenuBar(QMenuBar):
    menuTitlePressed = Signal(QAction)

    def mousePressEvent(self, e):
        act = self.actionAt(e.position().toPoint())
        if act:
            self.menuTitlePressed.emit(act)
        super().mousePressEvent(e)


class RibbonButton(QToolButton):
    def __init__(self, text: str, std_icon: QStyle.StandardPixmap, parent=None):
        super().__init__(parent)
        self.setText(text)
        self.setIcon(self.style().standardIcon(std_icon))
        self.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.setAutoRaise(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setIconSize(QSize(18, 18))
        self.setFixedSize(68, 58)
        self.setStyleSheet("""
            QToolButton { border: 1px solid transparent; border-radius: 12px; padding: 2px; }
            QToolButton:hover { background-color: #deeaf7; border: 1px solid #0078d7; }
        """)


class GutterShadow(QWidget):
    """페이지 내부(책등 방향) 인너 그림자"""
    def __init__(self, parent=None, side: str = "none"):
        super().__init__(parent)
        self.side = side  # "left" or "right" or "none"
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setFixedWidth(18)

    def set_side(self, side: str):
        self.side = side
        self.update()

    def paintEvent(self, e):
        if self.side == "none":
            return
        p = QPainter(self)
        w = self.width()
        h = self.height()
        grad_alpha = 55
        if self.side == "right":
            for x in range(w):
                a = int((x / max(1, w - 1)) * grad_alpha)
                p.fillRect(QRect(x, 0, 1, h), QColor(0, 0, 0, a))
        else:
            for x in range(w):
                a = int(((w - 1 - x) / max(1, w - 1)) * grad_alpha)
                p.fillRect(QRect(x, 0, 1, h), QColor(0, 0, 0, a))
        p.end()


class TextPageWidget(QFrame):
    """
    2페이지 0px 밀착을 위해:
    - 중앙에 shadow 번짐을 만들기 쉬운 drop shadow 효과 제거
    - 바깥쪽(outer)만 코너 그림자 직접 paint
    - 중앙 맞닿는 쪽 radius 0
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("BookPage")
        self.setFrameShape(QFrame.NoFrame)
        self.setMouseTracking(True)

        self._bg = QColor("#ffffff")

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        self.text = NoScrollTextBrowser()
        v.addWidget(self.text, 1)

        self.caption = QLabel("", self)
        self.caption.setFixedHeight(22)
        self.caption.setAlignment(Qt.AlignCenter)
        self.caption.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.caption.raise_()

        self._gutter = GutterShadow(self, side="none")
        self._gutter.hide()

        self._outer_edge = "right"
        self._edge_lift = 0.25
        self._edge_anim = QVariantAnimation(self)
        self._edge_anim.setDuration(140)
        self._edge_anim.setEasingCurve(QEasingCurve.OutCubic)
        self._edge_anim.valueChanged.connect(self._on_edge_anim)

        self.set_page_bg_color(QColor("#ffffff"))

    def set_outer_edge(self, side: str):
        self._outer_edge = "left" if side == "left" else "right"
        self.set_page_bg_color(self._bg)
        
        thickness = 8
        if self._outer_edge == "left":
            self.layout().setContentsMargins(thickness, 0, 0, 0)
        else:
            self.layout().setContentsMargins(0, 0, thickness, 0)
            
        self.update()

    def _on_edge_anim(self, v):
        self._edge_lift = float(v)
        self.update()

    def set_gutter_side(self, side: str):
        if side not in ("left", "right", "none"):
            side = "none"
        self._gutter.set_side(side)
        self._gutter.setVisible(side != "none")
        self._place_gutter()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._place_gutter()
        self._place_caption()

    def _place_gutter(self):
        if not self._gutter.isVisible():
            return
        w = self._gutter.width()
        if self._gutter.side == "right":
            self._gutter.setGeometry(self.width() - w, 0, w, self.height())
        elif self._gutter.side == "left":
            self._gutter.setGeometry(0, 0, w, self.height())

    def _place_caption(self):
        self.caption.setGeometry(0, self.height() - self.caption.height(), self.width(), self.caption.height())
        self.caption.raise_()

    def set_page_bg_color(self, c: QColor):
        self._bg = QColor(c)

        r = 12
        if getattr(self, "_outer_edge", "right") == "left":
            self.setStyleSheet(
                f"QFrame#BookPage {{ background:{c.name()}; "
                f"border-top-left-radius:{r}px; border-bottom-left-radius:{r}px; "
                f"border-top-right-radius:0px; border-bottom-right-radius:0px; }}"
            )
        else:
            self.setStyleSheet(
                f"QFrame#BookPage {{ background:{c.name()}; "
                f"border-top-right-radius:{r}px; border-bottom-right-radius:{r}px; "
                f"border-top-left-radius:0px; border-bottom-left-radius:0px; }}"
            )

        self.text.setStyleSheet(f"QTextBrowser {{ background:{c.name()}; border:none; }}")
        pal = self.text.viewport().palette()
        pal.setColor(self.text.viewport().backgroundRole(), c)
        self.text.viewport().setPalette(pal)
        self.text.viewport().setAutoFillBackground(True)

        # ✅ 페이지 번호(캡션) Bold 제거
        self.caption.setStyleSheet(
            f"QLabel {{ background:{c.name()}; color:#475569; font-size:18px; font-weight:400; padding:0 10px; }}"
        )

    def paintEvent(self, e):
        super().paintEvent(e)

        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing)

        w = self.width()
        h = self.height()

        corner_r = 34
        alpha = 80
        outer = getattr(self, "_outer_edge", "right")

        def corner_shadow(cx, cy, rect: QRect):
            g = QRadialGradient(QPoint(cx, cy), corner_r)
            g.setColorAt(0.0, QColor(0, 0, 0, alpha))
            g.setColorAt(1.0, QColor(0, 0, 0, 0))
            p.fillRect(rect, g)

        if outer == "left":
            corner_shadow(corner_r, corner_r, QRect(0, 0, corner_r * 2, corner_r * 2))
            corner_shadow(corner_r, h - corner_r, QRect(0, h - corner_r * 2, corner_r * 2, corner_r * 2))
        else:
            corner_shadow(w - corner_r, corner_r, QRect(w - corner_r * 2, 0, corner_r * 2, corner_r * 2))
            corner_shadow(w - corner_r, h - corner_r, QRect(w - corner_r * 2, h - corner_r * 2, corner_r * 2, corner_r * 2))

        # 2. Draw Book Thickness (Pages Edge)
        r = 12
        path = QPainterPath()
        path.setFillRule(Qt.WindingFill)
        path.addRoundedRect(0, 0, w, h, r, r)
        if outer == "left":
            path.addRect(r, 0, w - r, h)
        else:
            path.addRect(0, 0, w - r, h)
            
        p.setClipPath(path)

        thickness_w = 7
        dark_alpha = 50
        
        if outer == "left":
            grad = QLinearGradient(0, 0, thickness_w, 0)
            grad.setColorAt(0.0, QColor(0, 0, 0, dark_alpha))
            grad.setColorAt(1.0, QColor(0, 0, 0, 0))
            p.fillRect(0, 0, thickness_w, h, grad)
            
            p.setPen(QColor(0, 0, 0, 20))
            for i in range(1, thickness_w, 2):
                p.drawLine(i, 0, i, h)
                
            p.setPen(QColor(255, 255, 255, 80))
            for i in range(2, thickness_w, 2):
                p.drawLine(i, 0, i, h)
        else:
            grad = QLinearGradient(w - thickness_w, 0, w, 0)
            grad.setColorAt(0.0, QColor(0, 0, 0, 0))
            grad.setColorAt(1.0, QColor(0, 0, 0, dark_alpha))
            p.fillRect(w - thickness_w, 0, thickness_w, h, grad)
            
            p.setPen(QColor(0, 0, 0, 20))
            for i in range(1, thickness_w, 2):
                x = w - i
                p.drawLine(x, 0, x, h)
                
            p.setPen(QColor(255, 255, 255, 80))
            for i in range(2, thickness_w, 2):
                x = w - i
                p.drawLine(x, 0, x, h)

        p.end()


class EPUBReader(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = self.load_settings()
        self.setWindowTitle("EPUB Reader V2.14 (EPUB/TXT) - 2026.06.28")

        w, h = self.settings.get("window_size", (1000, 900))
        self.resize(int(w), int(h))

        self.view_profile = None
        self.doc_type = None
        self.path = None
        self.page_idx = 0

        self.view_mode = int(self.settings.get("view_mode", 2))  # 1 or 2
        self.thumb_visible = bool(self.settings.get("thumb_visible", True))
        self.theme_name = self.settings.get("theme", "종이")

        self.paper_name = self.settings.get("paper", "전체보기(창맞춤)")
        if self.paper_name not in PAPER_PRESETS:
            self.paper_name = "전체보기(창맞춤)"

        self.highlights = list(self.settings.get("highlights", []))
        self.bookmarks = list(self.settings.get("bookmarks", []))
        self._bookmark_pages = set(
            int(b.get("page", -1)) for b in self.bookmarks if isinstance(b, dict)
        )

        self.text_layout = TextLayoutSettings(
            font_family=self.settings.get("text_font_family", "맑은 고딕"),
            font_size=int(self.settings.get("text_font_size", 13)),
            bold=bool(self.settings.get("text_bold", False)),
            line_spacing=float(self.settings.get("line_spacing", 1.25)),
            para_spacing=int(self.settings.get("para_spacing", 8)),
            indent=float(self.settings.get("text_indent", 1.5)),
            margin_lr=int(self.settings.get("margin_lr", 40)),
            margin_tb=int(self.settings.get("margin_tb", 40)),
        )
        self.text_pager = TextPager(self.text_layout, DEFAULT_TEXT_PAPER)

        self._bg = QColor(THEMES.get(self.theme_name, THEMES["종이"]))
        self._fg = QColor("#e5e7eb") if self.theme_name == "블랙" else QColor("#111111")
        self.text_pager.set_colors(self._fg.name(), self._bg.name())

        self.session_start = time.time()
        self.total_time = int(self.settings.get("total_time", 0))

        self.text_thumb_cache = OrderedDict()
        self._thumb_placeholder_icon = self._build_placeholder_icon()

        self._render_debounce = QTimer(self)
        self._render_debounce.setSingleShot(True)
        self._render_debounce.timeout.connect(self.render_view)

        self._wheel_accum_angle = 0
        self._wheel_accum_pixel = 0

        self._thumb_timer = QTimer(self)
        self._thumb_timer.setSingleShot(True)
        self._thumb_queue = deque()
        self._thumb_pending = set()
        self._thumb_timer.timeout.connect(self._thumb_process_batch)

        self._thumb_debounce = QTimer(self)
        self._thumb_debounce.setSingleShot(True)
        self._thumb_debounce.timeout.connect(self._thumb_prepare_visible_queue)

        self._thumb_render_timer = QTimer(self)
        self._thumb_render_timer.setSingleShot(True)
        self._thumb_render_timer.timeout.connect(self.render_thumbnails)

        self._layout_change_timer = QTimer(self)
        self._layout_change_timer.setSingleShot(True)
        self._layout_change_timer.timeout.connect(self._apply_layout_change)

        self._save_settings_timer = QTimer(self)
        self._save_settings_timer.setSingleShot(True)
        self._save_settings_timer.timeout.connect(self.save_settings)

        self._thumb_scroll_anim = None

        self._search_term = ""
        self._search_hits = []
        self._search_pos = -1

        self.menu_activated_once = False
        self._last_rendered_page_idx = None

        self.build_ui()
        self.showMaximized()
        self.apply_ui_font(int(self.settings.get("ui_font_size", 12)))
        self.apply_theme()
        self.install_shortcuts()
        self.render_placeholder()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick_time)
        self.timer.start(1000)

        # 세션 복원/렌더 (중복 페이징 방지를 위해 단일 호출로 통합)
        QTimer.singleShot(0, self.restore_last_session)

    # ---------- Settings ----------
    def load_settings(self) -> Dict[str, Any]:
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {
                "ui_font_size": 12,
                "view_mode": 2,
                "thumb_visible": True,
                "paper": "전체보기(창맞춤)",
                "theme": "종이",
                "recent": [],
                "total_time": 0,
                "text_font_family": "맑은 고딕",
                "text_font_size": 13,
                "text_bold": False,
                "line_spacing": 1.25,
                "para_spacing": 8,
                "text_indent": 1.5,
                "margin_lr": 40,
                "margin_tb": 40,
                "last_path": "",
                "last_page_idx": 0,
                "highlights": [],
                "bookmarks": [],
                "window_size": (1300, 900),
            }

    def save_settings(self):
        self.settings["ui_font_size"] = int(QApplication.font().pointSize())
        self.settings["view_mode"] = int(self.view_mode)
        self.settings["thumb_visible"] = bool(self.thumb_visible)
        self.settings["paper"] = self.paper_name
        self.settings["theme"] = self.theme_name
        self.settings["total_time"] = int(self.total_time)

        self.settings["text_font_family"] = self.text_layout.font_family
        self.settings["text_font_size"] = int(self.text_layout.font_size)
        self.settings["text_bold"] = bool(self.text_layout.bold)
        self.settings["line_spacing"] = float(self.text_layout.line_spacing)
        self.settings["para_spacing"] = int(self.text_layout.para_spacing)
        self.settings["text_indent"] = float(self.text_layout.indent)
        self.settings["margin_lr"] = int(self.text_layout.margin_lr)
        self.settings["margin_tb"] = int(self.text_layout.margin_tb)

        self.settings["last_path"] = self.path or ""
        self.settings["last_page_idx"] = int(self.page_idx)

        self.settings["highlights"] = self.highlights
        self.settings["bookmarks"] = self.bookmarks
        self.settings["window_size"] = (int(self.width()), int(self.height()))

        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.settings, f, ensure_ascii=False, indent=2)

    # ---------- Helpers ----------
    def log(self, msg: str):
        self.sb_log.setText("상태: " + msg)

    def _start_load_ui(self):
        try:
            self._load_cancelled = False
            self.sb_progress.setValue(0)
            self.sb_progress.setVisible(True)
            self.sb_cancel.setVisible(True)
            self.sb_cancel.setEnabled(True)
            QApplication.processEvents()
        except Exception:
            pass

    def _stop_load_ui(self):
        try:
            self.sb_progress.setVisible(False)
            self.sb_cancel.setVisible(False)
            self.sb_progress.setValue(0)
            QApplication.processEvents()
        except Exception:
            pass

    def _on_cancel_load(self):
        try:
            self._load_cancelled = True
            self.sb_cancel.setEnabled(False)
            self.log('취소 요청됨')
            QApplication.processEvents()
        except Exception:
            pass

    def _file_load_progress(self, percent, msg=None):
        try:
            if isinstance(percent, int):
                self.sb_progress.setValue(percent)
            if msg:
                self.sb_log.setText('상태: ' + str(msg))
            QApplication.processEvents()
            # 취소 체크
            if getattr(self, '_load_cancelled', False):
                raise RuntimeError('로드 취소')
        except Exception:
            # 전파를 위해 재발생
            raise

    def _paging_progress(self, percent, msg=None):
        try:
            if isinstance(percent, int):
                self.sb_progress.setValue(percent)
            if msg:
                self.sb_log.setText('상태: ' + str(msg))
            QApplication.processEvents()
            if getattr(self, '_load_cancelled', False):
                raise RuntimeError('페이징 취소')
        except Exception:
            raise

    def clear_render_cache(self):
        self.text_thumb_cache.clear()

    def schedule_render(self):
        self._render_debounce.start(30)

    def schedule_thumb_update(self):
        self._thumb_debounce.start(35)

    def _remember_position(self):
        self.settings["last_path"] = self.path or ""
        self.settings["last_page_idx"] = int(self.page_idx)
        if self.path:
            file_pages = self.settings.setdefault("file_pages", {})
            file_pages[self.path] = int(self.page_idx)
        self.save_settings()

    def restore_last_session(self):
        last_path = self.settings.get("last_path", "")
        last_idx = int(self.settings.get("last_page_idx", 0) or 0)
        if last_path and isinstance(last_path, str) and os.path.exists(last_path):
            try:
                self.open_path(last_path, restore_page_idx=last_idx, silent=True)
                self.log("이전 세션 복원 완료")
            except Exception:
                self.render_placeholder()

    def _force_fit_to_window_initial(self):
        show_right = (self.view_mode == 2)
        self._layout_pages_only(show_right)
        if self.view_profile is None:
            self.render_placeholder()
        else:
            self.render_view()

    def _build_placeholder_icon(self) -> QIcon:
        pm = QPixmap(THUMB_ITEM_SIZE)
        pm.fill(QColor("#d1d5db"))
        painter = QPainter(pm)
        painter.setPen(QColor("#6b7280"))
        painter.drawRect(6, 6, THUMB_ITEM_SIZE.width() - 12, THUMB_ITEM_SIZE.height() - 12)
        painter.drawText(pm.rect(), Qt.AlignCenter, "…")
        painter.end()
        return QIcon(pm)

    # ---------- Theme ----------
    def apply_ui_font(self, pt: int):
        pt = clamp(int(pt), 8, 40)
        f = safe_ui_font()
        f.setPointSize(pt)
        QApplication.setFont(f)
        self.menuBar().setFont(f)
        self.statusBar().setFont(f)
        # QApplication.setFont() 이후 모든 위젯의 sizeHint가 재계산되어
        # 창 최소 크기(mintrack)가 화면 높이를 초과하는 문제를 방지한다.
        # 논리적 최솟값을 명시 고정하면 Qt가 이 값을 기준으로 제한한다.
        self.setMinimumSize(600, 400)
        # 리본 높이가 폰트 변경으로 팽창하지 않도록 재적용
        if hasattr(self, 'ribbon_stack'):
            self.ribbon_stack.setFixedHeight(76)
        self.save_settings()
        self.schedule_render()

    def apply_theme(self):
        self._bg = QColor(THEMES.get(self.theme_name, THEMES["종이"]))
        self._fg = QColor("#e5e7eb") if self.theme_name == "블랙" else QColor("#111111")
        self.text_pager.set_colors(self._fg.name(), self._bg.name())

        self.viewer.setStyleSheet(f"background-color:{CANVAS_BG};")
        self.page_left.set_page_bg_color(self._bg)
        self.page_right.set_page_bg_color(self._bg)

        self.clear_render_cache()
        self.render_thumbnails()
        self.schedule_render()

    # ---------- Paper ----------
    def set_paper(self, name: str):
        if name not in PAPER_PRESETS:
            name = "전체보기(창맞춤)"
        self.paper_name = name

        # 리본 용지 버튼 체크 상태 갱신
        try:
            for pname, btn in getattr(self, '_paper_btns', {}).items():
                btn.setChecked(pname == self.paper_name)
        except Exception:
            pass

        try:
            for a in getattr(self, "_paper_actions", []):
                a.setChecked(a.text() == self.paper_name)
        except Exception:
            pass

        self.save_settings()
        self.clear_render_cache()
        self.render_thumbnails()
        self.schedule_render()
        self.log(f"용지 크기: {self.paper_name}")

    # ---------- Font size ribbon helper ----------
    def set_text_font_size(self, pt: int):
        pt = clamp(int(pt), 6, 72)
        if int(self.text_layout.font_size) == pt:
            return
        self.text_layout.font_size = int(pt)
        try:
            if hasattr(self, "sp_font_size") and self.sp_font_size.value() != pt:
                self.sp_font_size.blockSignals(True)
                self.sp_font_size.setValue(pt)
                self.sp_font_size.blockSignals(False)
        except Exception:
            pass
        self._layout_change_timer.start(300)

    def set_ui_font_size(self, pt: int):
        pt = clamp(int(pt), 8, 30)
        self.settings["ui_font_size"] = pt
        try:
            if hasattr(self, "sp_ui_font") and self.sp_ui_font.value() != pt:
                self.sp_ui_font.blockSignals(True)
                self.sp_ui_font.setValue(pt)
                self.sp_ui_font.blockSignals(False)
        except Exception:
            pass
        self.apply_ui_font(pt)
        self.save_settings()

    # ---------- Menu UX ----------
    def on_any_menu_opened(self):
        self.menu_activated_once = True

    def on_menubar_hovered(self, action: QAction):
        if not self.menu_activated_once:
            return
        if not action:
            return
        m = action.menu()
        if m and not m.isVisible():
            pos = self.menuBar().mapToGlobal(self.menuBar().actionGeometry(action).bottomLeft())
            m.popup(pos)

    def on_menu_title_pressed(self, action: QAction):
        mapping = {
            "파일": 0,
            "용지 크기": 1,
            "보기": 2,
            "폰트": 3,
            "검색/표시": 4,
            "용지색상": 5,
        }
        t = action.text()
        if t in mapping:
            self.ribbon_stack.setCurrentIndex(mapping[t])

    # ---------- Events ----------
    def eventFilter(self, source, event):
        if event.type() == QEvent.Wheel and source in (
            self.viewer, self.spread,
            self.page_left, self.page_left.text,
            self.page_right, self.page_right.text,
        ):
            dy_pixel = 0
            dy_angle = 0
            try:
                if hasattr(event, "pixelDelta") and not event.pixelDelta().isNull():
                    dy_pixel = int(event.pixelDelta().y())
                dy_angle = int(event.angleDelta().y())
            except Exception:
                pass

            if dy_pixel != 0:
                self._wheel_accum_pixel += dy_pixel
                if abs(self._wheel_accum_pixel) >= WHEEL_PIXEL_THRESHOLD:
                    if self._wheel_accum_pixel > 0:
                        self.prev_page()
                    else:
                        self.next_page()
                    self._wheel_accum_pixel = 0
                    self._wheel_accum_angle = 0
            else:
                self._wheel_accum_angle += dy_angle
                if abs(self._wheel_accum_angle) >= WHEEL_ANGLE_THRESHOLD:
                    if self._wheel_accum_angle > 0:
                        self.prev_page()
                    else:
                        self.next_page()
                    self._wheel_accum_angle = 0
                    self._wheel_accum_pixel = 0

            return True

        if source == self.thumb.viewport() and event.type() in (QEvent.Resize, QEvent.Show):
            self.schedule_thumb_update()

        return super().eventFilter(source, event)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.schedule_render()

    # ---------- Recent ----------
    def refresh_recent_menu(self):
        self.menu_recent.clear()
        recent = self.settings.get("recent", [])
        if not recent:
            a = QAction("(없음)", self)
            a.setEnabled(False)
            self.menu_recent.addAction(a)
            return
        for pth in recent:
            a = QAction(pth, self)
            a.triggered.connect(lambda checked=False, path=pth: self.open_path(path))
            self.menu_recent.addAction(a)

    def push_recent(self, path: str):
        recent = self.settings.get("recent", [])
        if path in recent:
            recent.remove(path)
        recent.insert(0, path)
        self.settings["recent"] = recent[:12]
        self.refresh_recent_menu()
        self.save_settings()

    def open_recent_dialog(self):
        recent = self.settings.get("recent", [])
        if not recent:
            QMessageBox.information(self, "최근", "최근 파일이 없습니다.")
            return
        item, ok = QInputDialog.getItem(self, "최근 파일", "선택:", recent, 0, False)
        if ok and item:
            self.open_path(item)

    # ---------- File open ----------
    def open_file(self):
        filt = "Documents (*.epub *.txt)"
        path, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", filt)
        if path:
            self.open_path(path)

    def reload_with_encoding(self, encoding: str):
        """현재 열린 TXT 파일을 지정한 인코딩으로 다시 로드한다."""
        if not getattr(self, 'path', None):
            return
        ext = os.path.splitext(self.path)[1].lower()
        if ext != '.txt':
            QMessageBox.information(self, "인코딩 재시도", "TXT 파일에서만 사용할 수 있습니다.")
            return
        try:
            self.text_pager.load_txt(self.path, encoding=encoding)
            if self.text_pager.page_count() <= 0:
                raise ValueError("페이지 없음")
            self.page_idx = 0
            self.clear_render_cache()
            self.render_thumbnails()
            self.schedule_render()
            enc_label = self.text_pager.last_encoding
            self.log(f"인코딩 재시도: {enc_label}")
            self.sb_encoding.setText(f"[{enc_label}]")
        except Exception as e:
            QMessageBox.warning(self, "인코딩 오류", f"{encoding} 으로 열기 실패:\n{e}")

    def open_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택")
        if not folder:
            return
        files = []
        for root, _, names in os.walk(folder):
            for n in names:
                ext = os.path.splitext(n)[1].lower()
                if ext in SUPPORTED_OPEN_EXT:
                    files.append(os.path.join(root, n))
        files.sort()
        if not files:
            QMessageBox.information(self, "폴더", "지원하는 문서가 없습니다.")
            return
        item, ok = QInputDialog.getItem(self, "폴더 열기", "파일 선택:", files, 0, False)
        if ok and item:
            self.open_path(item)

    def open_path(self, path: str, restore_page_idx: Optional[int] = None, silent: bool = False):
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".txt", ".epub"):
            if not silent:
                QMessageBox.information(self, "지원 형식", "이 버전은 TXT와 EPUB만 지원합니다.")
            return

        try:
            # 기존 텍스트를 미리 클리어하여 open_path 초기 레이아웃 계산 시 불필요한 중복 페이징 방지
            self.text_pager.text = ""
            self.text_pager.pages.clear()

            show_right = (self.view_mode == 2)
            self._layout_pages_only(show_right)

            self.path = path
            self.doc_type = ext[1:] if ext.startswith(".") else ext

            # 시작 UI
            self._start_load_ui()
            self.text_pager.progress_callback = self._paging_progress
            try:
                if ext == ".txt":
                    self.view_profile = "reflow_txt"
                    self.text_pager.set_layout(self.text_layout)
                    self.text_pager.load_txt(path)
                    if self.text_pager.page_count() <= 0:
                        raise ValueError("TXT에서 텍스트를 읽었으나 페이지가 생성되지 않았습니다.")
                    # 인코딩 상태 표시
                    enc_label = getattr(self.text_pager, 'last_encoding', '')
                    if hasattr(self, 'sb_encoding'):
                        self.sb_encoding.setText(f"[{enc_label}]")
                else:
                    self.view_profile = "reflow_epub"
                    text = extract_text_from_epub(path, progress_callback=self._file_load_progress)
                    if not text.strip():
                        raise ValueError("EPUB에서 텍스트를 추출하지 못했습니다. (container.xml/opf/spine 또는 HTML 구조 확인)")
                    self.text_pager.set_layout(self.text_layout)
                    self.text_pager.load_text(text)
                    if self.text_pager.page_count() <= 0:
                        raise ValueError("EPUB 텍스트는 추출되었으나 페이지가 생성되지 않았습니다. (용지/여백/폰트 설정 확인)")
                    if hasattr(self, 'sb_encoding'):
                        self.sb_encoding.setText("")
            finally:
                # 종료 UI
                self.text_pager.progress_callback = None
                try:
                    self._stop_load_ui()
                except Exception:
                    pass

            if restore_page_idx is None:
                file_pages = self.settings.get("file_pages", {})
                saved_idx = file_pages.get(path)
                if saved_idx is not None:
                    self.page_idx = self._align_left_page(max(0, int(saved_idx)))
                else:
                    self.page_idx = 0
            else:
                self.page_idx = self._align_left_page(max(0, int(restore_page_idx)))

            self._last_rendered_page_idx = None

            self.push_recent(path)
            self.clear_render_cache()
            self.render_thumbnails()

            self.render_view()
            self._remember_position()
            self.log(f"열기 완료: {os.path.basename(path)} ({self.doc_type.upper()})")
        except Exception as e:
            if not silent:
                tb = traceback.format_exc()
                QMessageBox.warning(self, "열기 실패", f"{type(e).__name__}: {e}\n\n{tb}")
            self.view_profile = None
            self.render_placeholder()

    # ---------- View / Page ----------
    def page_count(self) -> int:
        if self.view_profile is None:
            return 0
        return self.text_pager.page_count()

    def _align_left_page(self, idx: int) -> int:
        if self.view_mode == 2:
            return idx if (idx % 2 == 0) else max(0, idx - 1)
        return idx

    def go_to_page(self, idx: int):
        if idx < 0:
            return
        if idx < self.page_count():
            self.page_idx = self._align_left_page(idx)
            self._remember_position()
            self.schedule_render()
            self._thumb_focus_current(animated=True)

    def next_page(self):
        step = self.view_mode
        if self.page_idx + step < self.page_count():
            self.page_idx = self._align_left_page(self.page_idx + step)
            self._remember_position()
            self.schedule_render()
            self._thumb_focus_current(animated=False)

    def prev_page(self):
        step = self.view_mode
        if self.page_idx - step >= 0:
            self.page_idx = self._align_left_page(self.page_idx - step)
            self._remember_position()
            self.schedule_render()
            self._thumb_focus_current(animated=False)

    # ---------- Search ----------
    def search_text(self):
        if self.view_profile is None:
            return
        term, ok = QInputDialog.getText(self, "검색", "검색어:")
        if not ok or not term.strip():
            return
        term = term.strip()
        self._search_term = term
        self._search_hits = [i for i, chunk in enumerate(self.text_pager.pages) if term in chunk]
        if not self._search_hits:
            self.log(f"검색 실패: {term}")
            self.schedule_render()
            return
        self._search_pos = 0
        self.page_idx = self._align_left_page(self._search_hits[0])
        self._remember_position()
        self.schedule_render()
        self._thumb_focus_current(animated=True)
        self.log(f"검색: '{term}' (1/{len(self._search_hits)})")

    def find_next(self):
        if not self._search_term or not self._search_hits:
            self.search_text()
            return
        self._search_pos = (self._search_pos + 1) % len(self._search_hits)
        self.page_idx = self._align_left_page(self._search_hits[self._search_pos])
        self._remember_position()
        self.schedule_render()
        self._thumb_focus_current(animated=True)
        self.log(f"다음: '{self._search_term}' ({self._search_pos+1}/{len(self._search_hits)})")

    def find_prev(self):
        if not self._search_term or not self._search_hits:
            self.search_text()
            return
        self._search_pos = (self._search_pos - 1) % len(self._search_hits)
        self.page_idx = self._align_left_page(self._search_hits[self._search_pos])
        self._remember_position()
        self.schedule_render()
        self._thumb_focus_current(animated=True)
        self.log(f"이전: '{self._search_term}' ({self._search_pos+1}/{len(self._search_hits)})")

    # ---------- Highlight & Bookmark ----------
    def _focused_text_widget(self) -> QTextBrowser:
        for w in (self.page_left.text, self.page_right.text):
            try:
                if w.hasFocus() and w.textCursor().hasSelection():
                    return w
            except Exception:
                pass
        for w in (self.page_left.text, self.page_right.text):
            if w.hasFocus():
                return w
        return self.page_left.text

    def add_user_highlight(self):
        if self.view_profile is None:
            return
        w = self._focused_text_widget()
        cur = w.textCursor()
        if not cur.hasSelection():
            QMessageBox.information(self, "하이라이트", "먼저 텍스트를 선택한 뒤 하이라이트를 추가하세요.")
            return
        text = cur.selectedText().replace("\u2029", "\n").strip()
        if len(text) < 2:
            return
        for h in self.highlights:
            if isinstance(h, dict) and h.get("text") == text:
                self.log("이미 하이라이트에 등록된 텍스트입니다.")
                self.schedule_render()
                return

        self.highlights.append({"text": text, "color": DEFAULT_USER_HIGHLIGHT_COLOR})
        self.save_settings()
        self.schedule_render()
        self.log("하이라이트 추가됨")

    def clear_user_highlights(self):
        self.highlights = []
        self.save_settings()
        self.schedule_render()
        self.log("하이라이트 전체 삭제")

    def add_bookmark(self):
        if self.view_profile is None:
            return
        page = int(self.page_idx)
        label, ok = QInputDialog.getText(self, "책갈피", "메모(선택):")
        if not ok:
            return
        label = (label or "").strip()
        updated = False
        for b in self.bookmarks:
            if isinstance(b, dict) and int(b.get("page", -1)) == page:
                b["label"] = label
                updated = True
                break
        if not updated:
            self.bookmarks.append({"page": page, "label": label})

        self._bookmark_pages = set(int(b.get("page", -1)) for b in self.bookmarks if isinstance(b, dict))
        self.save_settings()
        self._refresh_bookmark_stars()
        self.schedule_thumb_update()
        self.schedule_render()
        self.log("책갈피 추가/갱신")

    def goto_bookmark(self):
        if not self.bookmarks:
            QMessageBox.information(self, "책갈피", "책갈피가 없습니다.")
            return
        items = []
        for b in self.bookmarks:
            if not isinstance(b, dict):
                continue
            p = int(b.get("page", 0))
            label = (b.get("label") or "").strip()
            disp = f"p.{p+1}" + (f" - {label}" if label else "")
            items.append(disp)
        if not items:
            QMessageBox.information(self, "책갈피", "책갈피가 없습니다.")
            return
        sel, ok = QInputDialog.getItem(self, "책갈피 목록", "이동:", items, 0, False)
        if ok and sel:
            m = re.match(r"p\.(\d+)", sel)
            if m:
                p = int(m.group(1)) - 1
                self.go_to_page(max(0, p))

    def remove_bookmark_current(self):
        if self.view_profile is None or not self.bookmarks:
            return
        page = int(self.page_idx)
        before = len(self.bookmarks)
        self.bookmarks = [b for b in self.bookmarks if not (isinstance(b, dict) and int(b.get("page", -1)) == page)]
        if len(self.bookmarks) != before:
            self._bookmark_pages = set(int(b.get("page", -1)) for b in self.bookmarks if isinstance(b, dict))
            self.save_settings()
            self._refresh_bookmark_stars()
            self.schedule_thumb_update()
            self.schedule_render()
            self.log("현재 페이지 책갈피 삭제")
        else:
            self.log("현재 페이지에 책갈피가 없습니다.")

    def _refresh_bookmark_stars(self):
        if self.thumb.count() <= 0:
            return
        for i in range(self.thumb.count()):
            it = self.thumb.item(i)
            if not it:
                continue
            star = "★ " if i in self._bookmark_pages else ""
            it.setText(f"{star}{i+1}")

    def _normalize_search_term(self, term: str) -> str:
        if not term:
            return ""
        result = []
        for i, ch in enumerate(term):
            result.append(ch)
            if "\uAC00" <= ch <= "\uD7A3":
                if i + 1 < len(term):
                    nxt = term[i + 1]
                    if "\uAC00" <= nxt <= "\uD7A3":
                        result.append("\u200b")
        return "".join(result)

    # ---------- Highlight rendering ----------
    def _collect_search_selections(self, widget: QTextBrowser):
        term = self._normalize_search_term(self._search_term or "")
        if not term:
            return []
        doc = widget.document()
        cursor = QTextCursor(doc)
        selections = []
        fmt = QTextCharFormat()
        fmt.setBackground(QColor(DEFAULT_SEARCH_HIGHLIGHT_COLOR))
        while True:
            cursor = doc.find(term, cursor)
            if cursor.isNull():
                break
            sel = QTextEdit.ExtraSelection()
            sel.cursor = cursor
            sel.format = fmt
            selections.append(sel)
        return selections

    def _collect_user_highlight_selections(self, widget: QTextBrowser):
        if not self.highlights:
            return []
        doc = widget.document()
        selections = []
        for h in self.highlights:
            if not isinstance(h, dict):
                continue
            term = self._normalize_search_term((h.get("text") or "").strip())
            if not term:
                continue
            color = h.get("color") or DEFAULT_USER_HIGHLIGHT_COLOR
            cursor = QTextCursor(doc)
            fmt = QTextCharFormat()
            fmt.setBackground(QColor(color))
            while True:
                cursor = doc.find(term, cursor)
                if cursor.isNull():
                    break
                sel = QTextEdit.ExtraSelection()
                sel.cursor = cursor
                sel.format = fmt
                selections.append(sel)
        return selections

    def _apply_all_highlights(self, widget: QTextBrowser):
        sels = []
        sels.extend(self._collect_user_highlight_selections(widget))
        sels.extend(self._collect_search_selections(widget))
        widget.setExtraSelections(sels)

    # ---------- Actions ----------
    def set_theme(self, name: str):
        self.theme_name = name
        # 리송 테마 버튼 체크 상태 갱신
        for tname, btn in getattr(self, '_theme_btns', {}).items():
            btn.setChecked(tname == name)
        self.save_settings()
        self.apply_theme()
        self.log(f"용지색상: {name}")

    def toggle_thumbnails(self, checked: bool):
        self.thumb_visible = bool(checked)
        self.thumb.setVisible(self.thumb_visible)
        self.thumb_btn.setChecked(self.thumb_visible)
        self.act_thumb.setChecked(self.thumb_visible)

        if self.thumb_visible:
            self.thumb.setFixedWidth(260)
            self.splitter.setSizes([260, max(1, self.width() - 260)])
            self.schedule_thumb_update()
        else:
            self.splitter.setSizes([0, max(1, self.width())])

        self.save_settings()
        self.schedule_render()

    def set_view_mode(self, mode: int):
        self.view_mode = 1 if mode == 1 else 2
        self.page_idx = self._align_left_page(self.page_idx)
        self.save_settings()
        self.schedule_render()
        self._thumb_focus_current(animated=False)

    def open_font_dialog(self):
        cur = QFont(self.text_layout.font_family, self.text_layout.font_size)
        cur.setBold(self.text_layout.bold)
        ok, font = QFontDialog.getFont(cur, self, "폰트 선택")
        if ok:
            fam = font.family()
            if fam in BAD_FONTS:
                self.log("선택한 폰트가 제한됨(BAD_FONTS).")
                return
            self.text_layout.font_family = fam
            if font.pointSize() > 0:
                self.text_layout.font_size = font.pointSize()
                self.set_text_font_size(self.text_layout.font_size)
            self.text_layout.bold = font.bold()
            self.bold_btn.setChecked(self.text_layout.bold)
            self._apply_layout_change()

    def set_text_bold(self, on: bool):
        self.text_layout.bold = bool(on)
        self.bold_btn.setChecked(self.text_layout.bold)
        self._layout_change_timer.start(300)

    def set_margin_lr(self, v: int):
        self.text_layout.margin_lr = int(v)
        self._apply_temporary_viewport_margins()
        self._layout_change_timer.start(300)

    def set_margin_tb(self, v: int):
        self.text_layout.margin_tb = int(v)
        self._apply_temporary_viewport_margins()
        self._layout_change_timer.start(300)

    def set_line_spacing(self, v: float):
        self.text_layout.line_spacing = float(v)
        self._layout_change_timer.start(300)

    def set_para_spacing(self, v: int):
        self.text_layout.para_spacing = int(v)
        self._layout_change_timer.start(300)

    def set_para_indent(self, v: float):
        self.text_layout.indent = float(v)
        self._layout_change_timer.start(300)

    def _apply_layout_change(self):
        self.text_pager.set_layout(self.text_layout)
        self.clear_render_cache()
        if self.thumb_visible:
            self._thumb_render_timer.start(150)
        self.schedule_render()
        self._save_settings_timer.start(1000)

    def _apply_temporary_viewport_margins(self):
        mlr = int(self.text_layout.margin_lr)
        mtb = int(self.text_layout.margin_tb)
        top_margin = 15 + mtb
        bottom_margin = 5 + mtb
        for w in (self.page_left.text, self.page_right.text):
            if w is not None:
                w.setViewportMargins(mlr, top_margin, mlr, bottom_margin)
                w.viewport().update()

    # ---------- Thumbnail lazy ----------
    def _thumb_cache_get(self, key):
        if key in self.text_thumb_cache:
            self.text_thumb_cache.move_to_end(key)
            return self.text_thumb_cache[key]
        return None

    def _thumb_cache_put(self, key, icon: QIcon):
        self.text_thumb_cache[key] = icon
        self.text_thumb_cache.move_to_end(key)
        while len(self.text_thumb_cache) > TEXT_THUMB_CACHE_MAX:
            self.text_thumb_cache.popitem(last=False)

    def render_thumbnails(self):
        if not self.thumb_visible:
            return
        self.thumb.clear()
        if self.view_profile is None:
            return

        total = self.text_pager.page_count()
        self.thumb.setUpdatesEnabled(False)

        for i in range(total):
            star = "★ " if i in self._bookmark_pages else ""
            item = QListWidgetItem(f"{star}{i+1}")
            item.setIcon(self._thumb_placeholder_icon)
            self.thumb.addItem(item)

        self.thumb.setUpdatesEnabled(True)

        self._thumb_focus_current(animated=False)
        self.schedule_thumb_update()

    def _thumb_focus_current(self, animated: bool):
        if not self.thumb_visible or self.thumb.count() <= 0:
            return
        row = int(self.page_idx)
        if row < 0 or row >= self.thumb.count():
            return

        if self.thumb.currentRow() != row:
            self.thumb.blockSignals(True)
            self.thumb.setCurrentRow(row)
            self.thumb.blockSignals(False)

        it = self.thumb.item(row)
        if not it:
            return

        if animated:
            self._animate_thumb_scroll_to_item(it)
        else:
            self.thumb.scrollToItem(it, QListWidget.PositionAtCenter)

        self.schedule_thumb_update()

    def _animate_thumb_scroll_to_item(self, item: QListWidgetItem):
        sb = self.thumb.verticalScrollBar()
        rect = self.thumb.visualItemRect(item)
        viewport_h = self.thumb.viewport().height()
        delta = rect.center().y() - (viewport_h // 2)
        target = clamp(sb.value() + delta, sb.minimum(), sb.maximum())

        old_val = sb.value()
        if self._thumb_scroll_anim:
            try:
                self._thumb_scroll_anim.stop()
            except Exception:
                pass
            self._thumb_scroll_anim = None

        anim = QPropertyAnimation(sb, b"value", self)
        anim.setDuration(THUMB_SCROLL_ANIM_MS)
        anim.setStartValue(old_val)
        anim.setEndValue(target)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        anim.start(QPropertyAnimation.DeleteWhenStopped)
        self._thumb_scroll_anim = anim

    def _thumb_prepare_visible_queue(self):
        if not self.thumb_visible or self.view_profile is None:
            return
        if self.thumb.count() <= 0:
            return

        self._thumb_timer.stop()
        self._thumb_queue.clear()
        self._thumb_pending.clear()

        top = self.thumb.indexAt(QPoint(5, 5)).row()
        bottom = self.thumb.indexAt(QPoint(5, self.thumb.viewport().height() - 5)).row()
        if top < 0:
            top = 0
        if bottom < 0:
            bottom = min(self.thumb.count() - 1, top + 10)

        start = max(0, top - THUMB_PREFETCH_PAD)
        end = min(self.thumb.count() - 1, bottom + THUMB_PREFETCH_PAD)

        c = int(self.page_idx)
        pstart = max(0, c - THUMB_FAST_CENTER_PAD)
        pend = min(self.thumb.count() - 1, c + THUMB_FAST_CENTER_PAD)

        priority = list(range(pstart, pend + 1))
        rest = [i for i in range(start, end + 1) if i < pstart or i > pend]
        order = priority + rest

        sig = self.text_pager.signature()
        for i in order:
            key = ("thumb", i, self.theme_name, sig)
            cached = self._thumb_cache_get(key)
            it = self.thumb.item(i)
            if not it:
                continue
            if cached is not None:
                if it.icon().cacheKey() != cached.cacheKey():
                    it.setIcon(cached)
                continue
            if i not in self._thumb_pending:
                self._thumb_queue.append(i)
                self._thumb_pending.add(i)

        if self._thumb_queue:
            self._thumb_timer.start(0)

    def _thumb_process_batch(self):
        if not self.thumb_visible or self.view_profile is None:
            return
        if not self._thumb_queue:
            return

        bg = self._bg
        sig = self.text_pager.signature()
        count = 0

        self.thumb.setUpdatesEnabled(False)
        try:
            while self._thumb_queue and count < THUMB_BATCH:
                i = self._thumb_queue.popleft()
                self._thumb_pending.discard(i)
                count += 1

                key = ("thumb", i, self.theme_name, sig)
                icon = self._thumb_cache_get(key)
                if icon is None:
                    icon = self.text_pager.render_thumb_icon(i, bg)
                    self._thumb_cache_put(key, icon)

                it = self.thumb.item(i)
                if it:
                    it.setIcon(icon)
        finally:
            self.thumb.setUpdatesEnabled(True)

        if self._thumb_queue:
            self._thumb_timer.start(0)

    # ---------- Layout sizing ----------
    def _available_view_area(self) -> QSize:
        w = max(320, self.viewer.width() - 12)
        h = max(320, self.viewer.height() - 12)
        return QSize(w, h)

    def _page_aspect_ratio(self) -> float:
        preset = PAPER_PRESETS.get(self.paper_name)
        if preset is None:
            return DEFAULT_TEXT_PAPER.width() / max(1, DEFAULT_TEXT_PAPER.height())
        w_mm, h_mm = preset
        return float(w_mm) / max(1.0, float(h_mm))

    def _compute_page_slot(self, avail: QSize, two_pages: bool) -> Tuple[int, int]:
        ratio = self._page_aspect_ratio()
        if two_pages:
            slot_w = max(240, avail.width() // 2)
            slot_h = max(240, avail.height())
        else:
            slot_w = max(240, avail.width())
            slot_h = max(240, avail.height())

        page_w = slot_w
        page_h = int(page_w / max(0.01, ratio))
        if page_h > slot_h:
            page_h = slot_h
            page_w = int(page_h * ratio)
        return int(page_w), int(page_h)

    def _ensure_fixed_sizes(self, page_w: int, page_h: int, show_right: bool):
        self.page_left.setFixedSize(page_w, page_h)
        if show_right:
            self.page_right.setFixedSize(page_w, page_h)

        inner_w = max(120, page_w)
        inner_h = max(120, page_h)
        self.page_left.text.setFixedSize(inner_w, inner_h)
        if show_right:
            self.page_right.text.setFixedSize(inner_w, inner_h)

    def _text_paper_for_view(self, viewport: QSize) -> QSize:
        preset = PAPER_PRESETS.get(self.paper_name)
        if preset is None:
            return QSize(max(220, viewport.width()), max(220, viewport.height()))
        w_mm, h_mm = preset
        ratio = float(w_mm) / max(1.0, float(h_mm))
        h = viewport.height()
        w = int(h * ratio)
        if w > viewport.width():
            w = viewport.width()
            h = int(w / max(0.01, ratio))
        return QSize(max(220, w), max(220, h))

    def _layout_pages_only(self, show_right: bool):
        self.page_right.setVisible(show_right)

        avail = self._available_view_area()
        page_w, page_h = self._compute_page_slot(avail, two_pages=show_right)
        self._ensure_fixed_sizes(page_w, page_h, show_right)

        page_area = self.page_left.text.size()
        self.text_pager.set_paper_px(self._text_paper_for_view(page_area))

        if show_right:
            self.page_left.set_gutter_side("none")
            self.page_right.set_gutter_side("none")
            self.page_left.set_outer_edge("left")
            self.page_right.set_outer_edge("right")
        else:
            self.page_left.set_gutter_side("none")
            self.page_right.set_gutter_side("none")
            self.page_left.set_outer_edge("right")

        # spine overlay: 중앙 seam에 겹치기(공간 0)
        if show_right:
            self.spread.layout().activate()  # ✅ 레이아웃 강제 업데이트로 위치 정확도 확보
            self.spine.setVisible(True)
            left_geo = self.page_left.geometry()
            cx = left_geo.right()  # ✅ 오버랩이면 +1 없이 맞추는 편이 자연스러움
            self.spine.setGeometry(
                cx - self.spine.width() // 2,
                left_geo.top(),
                self.spine.width(),
                left_geo.height()
            )
            self.spine.raise_()
        else:
            self.spine.setVisible(False)

    # ---------- Placeholder / Render ----------
    def render_placeholder(self):
        msg = (
            "파일 > 불러오기에서 TXT/EPUB를 열어주세요.\n\n"
            "- 썸네일 휠: 페이지 선택(1장=1, 2장=2)\n\n"
            "- 뷰어 휠: 페이지 넘김\n\n"
            "- Ctrl+F: 검색, F3/Shift+F3: 다음/이전\n\n"
            "- Ctrl+H: 선택 하이라이트 추가\n\n"
            "- Ctrl+D: 책갈피 추가, Ctrl+B: 책갈피 목록\n\n"
            "- 용지 크기 메뉴에서 비율 변경 가능"
        )
        self.view_profile = None

        show_right = (self.view_mode == 2)
        self._layout_pages_only(show_right)

        mlr = int(self.text_layout.margin_lr)
        mtb = int(self.text_layout.margin_tb)
        top_margin = 15 + mtb
        bottom_margin = 5 + mtb

        f = QFont(self.text_layout.font_family, self.text_layout.font_size)
        f.setBold(self.text_layout.bold)
        f.setStyleHint(QFont.SansSerif)

        opt = QTextOption()
        opt.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        opt.setAlignment(Qt.AlignJustify | Qt.AlignTop)

        for w in (self.page_left.text, self.page_right.text):
            w.setViewportMargins(mlr, top_margin, mlr, bottom_margin)
            w.clear()
            w.setFont(f)
            w.document().setDocumentMargin(0)
            w.document().setDefaultFont(f)
            w.document().setDefaultStyleSheet(self.text_pager.build_css())
            w.document().setDefaultTextOption(opt)
            w.setHtml(self.text_pager.build_html(msg))

        self.page_left.caption.setText("")
        self.page_right.caption.setText("")
        self.sb_page.setText("페이지: 0/0 (0%)")
        self.log("준비됨")

    def render_view(self):
        if self.view_profile is None:
            self.render_placeholder()
            return

        total = self.page_count()
        if total <= 0:
            self.sb_page.setText("페이지: 0/0 (0%)")
            self._layout_pages_only(self.view_mode == 2)
            return

        show_right = (self.view_mode == 2)
        self._layout_pages_only(show_right)

        if self.page_idx >= total:
            self.page_idx = self._align_left_page(max(0, total - 1))
            show_right = (self.view_mode == 2)
            self._layout_pages_only(show_right)

        mlr = int(self.text_layout.margin_lr)
        mtb = int(self.text_layout.margin_tb)
        top_margin = 15 + mtb
        bottom_margin = 5 + mtb

        def apply_text(widget: QTextBrowser, page_no: int):
            widget.setViewportMargins(mlr, top_margin, mlr, bottom_margin)

            f = QFont(self.text_layout.font_family, self.text_layout.font_size)
            f.setBold(self.text_layout.bold)
            f.setStyleHint(QFont.SansSerif)

            widget.setFont(f)
            widget.document().setDocumentMargin(0)
            widget.document().setDefaultFont(f)

            # ✅ 폰트/CSS 설정값이 본문에 100% 반영되도록 강제
            widget.document().setDefaultStyleSheet(self.text_pager.build_css())
            opt = QTextOption()
            opt.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
            opt.setAlignment(Qt.AlignJustify | Qt.AlignTop)
            widget.document().setDefaultTextOption(opt)

            widget.clear()
            conts = self.text_pager.page_continuations
            is_cont = conts[page_no] if page_no < len(conts) else False
            is_next_cont = conts[page_no + 1] if (page_no + 1) < len(conts) else False
            widget.setHtml(self.text_pager.build_html(self.text_pager.pages[page_no], is_continuation=is_cont, is_continued_to_next=is_next_cont))
            self._apply_all_highlights(widget)
            widget.viewport().update()

        apply_text(self.page_left.text, self.page_idx)

        if show_right and (self.page_idx + 1) < total:
            apply_text(self.page_right.text, self.page_idx + 1)
        else:
            self.page_right.text.clear()
            self.page_right.text.setHtml(self.text_pager.build_html(""))

        # 캡션
        if self.view_mode == 1:
            self.page_left.caption.setAlignment(Qt.AlignCenter)
            self.page_left.caption.setText(f"{self.page_idx + 1} / {total}")
            self.page_right.caption.setText("")
        else:
            self.page_left.caption.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.page_right.caption.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.page_left.caption.setText(f"{self.page_idx + 1}")
            self.page_right.caption.setText(f"{self.page_idx + 2}" if (self.page_idx + 1) < total else "")

        prog = int(((self.page_idx + 1) / max(1, total)) * 100)
        self.sb_page.setText(f"페이지: {self.page_idx + 1}/{total} ({prog}%)")

        # 썸네일 동기화
        if self.thumb_visible and self.thumb.count() > 0:
            if self.thumb.currentRow() != self.page_idx:
                self.thumb.blockSignals(True)
                self.thumb.setCurrentRow(self.page_idx)
                self.thumb.blockSignals(False)
            self.schedule_thumb_update()

        self._last_rendered_page_idx = self.page_idx

    # ---------- Shortcuts ----------
    def install_shortcuts(self):
        def add_shortcut(keys, fn):
            act = QAction(self)
            act.setShortcut(QKeySequence(keys) if isinstance(keys, str) else keys)
            act.triggered.connect(fn)
            self.addAction(act)

        add_shortcut(QKeySequence(Qt.Key_Right), self.next_page)
        add_shortcut(QKeySequence(Qt.Key_Left), self.prev_page)
        add_shortcut(QKeySequence(Qt.Key_PageDown), self.next_page)
        add_shortcut(QKeySequence(Qt.Key_PageUp), self.prev_page)
        add_shortcut(QKeySequence(Qt.Key_Home), lambda: self.go_to_page(0))
        add_shortcut(QKeySequence(Qt.Key_End), lambda: self.go_to_page(max(0, self.page_count() - 1)))

        act_find = QAction(self)
        act_find.setShortcut(QKeySequence.Find)
        act_find.triggered.connect(self.search_text)
        self.addAction(act_find)

        act_next = QAction(self)
        act_next.setShortcut(QKeySequence("F3"))
        act_next.triggered.connect(self.find_next)
        self.addAction(act_next)

        act_prev = QAction(self)
        act_prev.setShortcut(QKeySequence("Shift+F3"))
        act_prev.triggered.connect(self.find_prev)
        self.addAction(act_prev)

        add_shortcut(QKeySequence("Ctrl+H"), self.add_user_highlight)
        add_shortcut(QKeySequence("Ctrl+D"), self.add_bookmark)
        add_shortcut(QKeySequence("Ctrl+B"), self.goto_bookmark)

    # ---------- Time ----------
    def tick_time(self):
        session = int(time.time() - self.session_start)
        total = session + self.total_time
        self.sb_session.setText("독서시간: " + time.strftime("%H:%M:%S", time.gmtime(session)))
        self.sb_total.setText("누적: " + time.strftime("%H:%M:%S", time.gmtime(total)))

    def closeEvent(self, e):
        self.total_time += int(time.time() - self.session_start)
        self.save_settings()
        e.accept()

    # ---------- UI ----------
    def build_ui(self):
        self.build_menubar()
        self.build_ribbon()

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self.ribbon_stack)

        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setHandleWidth(0)  # ✅ 썸네일 토글 시 레이아웃 흔들림/틈 최소화
        root.addWidget(self.splitter, 1)

        # thumbnails
        self.thumb = ThumbListWidget()
        self.thumb.setFixedWidth(260)
        self.thumb.setIconSize(THUMB_ITEM_SIZE)
        self.thumb.currentRowChanged.connect(self.go_to_page)
        self.thumb.setVisible(self.thumb_visible)
        self.thumb.visibleRangeChanged.connect(self.schedule_thumb_update)
        self.thumb.viewport().installEventFilter(self)
        self.splitter.addWidget(self.thumb)

        # viewer
        self.viewer = QWidget()
        self.viewer.setObjectName("Viewer")
        vlayout = QVBoxLayout(self.viewer)
        vlayout.setContentsMargins(6, 6, 6, 6)
        vlayout.setSpacing(0)

        self.spread = QWidget()
        self.spread.setObjectName("Spread")
        hlay = QHBoxLayout(self.spread)
        hlay.setContentsMargins(0, 0, 0, 0)
        # ✅ 창 크기에 따라 여백이 변하지 않고 중앙에 모이도록 정렬 설정
        hlay.setAlignment(Qt.AlignCenter)
        hlay.setSpacing(0)

        self.page_left = TextPageWidget()
        self.page_right = TextPageWidget()

        hlay.addWidget(self.page_left)
        hlay.addWidget(self.page_right)
        vlayout.addWidget(self.spread, 1)
        self.splitter.addWidget(self.viewer)

        # ✅ spine overlay 위젯 생성(오버레이 방식)
        self.spine = QWidget(self.spread)
        self.spine.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.spine.setFixedWidth(24)  # 책등 두께
        self.spine.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 rgba(0,0,0,0),
                    stop:0.35 rgba(0,0,0,35),
                    stop:0.5 rgba(0,0,0,55),
                    stop:0.65 rgba(0,0,0,35),
                    stop:1 rgba(0,0,0,0)
                );
            }
        """)
        self.spine.hide()

        if self.thumb_visible:
            self.splitter.setSizes([260, max(1, self.width() - 260)])
        else:
            self.splitter.setSizes([0, max(1, self.width())])

        for w in (self.viewer, self.spread, self.page_left, self.page_left.text, self.page_right, self.page_right.text):
            w.installEventFilter(self)

        # status
        sb = self.statusBar()
        self.sb_log = QLabel("상태: 준비됨")
        self.sb_page = QLabel("페이지: 0/0 (0%)")
        self.sb_session = QLabel("독서시간: 00:00:00")
        self.sb_total = QLabel("누적: 00:00:00")
        # 진행바와 취소 버튼
        self.sb_progress = QProgressBar()
        self.sb_progress.setMaximumWidth(180)
        self.sb_progress.setVisible(False)
        self.sb_cancel = QPushButton("취소")
        self.sb_cancel.setVisible(False)
        self.sb_cancel.clicked.connect(lambda: getattr(self, '_on_cancel_load', lambda: None)())

        sb.addWidget(self.sb_log, 1)
        sb.addPermanentWidget(self.sb_progress)
        sb.addPermanentWidget(self.sb_cancel)
        sb.addPermanentWidget(self.sb_page)
        sb.addPermanentWidget(self.sb_session)
        sb.addPermanentWidget(self.sb_total)

    def build_menubar(self):
        mb = SmartMenuBar()
        mb.setNativeMenuBar(False)
        mb.setMouseTracking(True)
        mb.setLayoutDirection(Qt.LeftToRight)
        # QMenuBar 및 item 스타일 지정을 통해 리본 메뉴와의 경계선을 선명하게 유지
        mb.setStyleSheet("""
            QMenuBar {
                background-color: #f8f9fa;
                border-bottom: 1px solid #d0d0d0;
            }
            QMenuBar::item {
                padding: 6px 14px;
                background: transparent;
            }
            QMenuBar::item:selected {
                background: #d0e4f7;
                border-radius: 4px;
            }
            QMenuBar::item:pressed {
                background: #b3d0f0;
                border-radius: 4px;
            }
        """)
        self.setMenuBar(mb)

        self.menu_file = RightHintMenu("파일", mb)
        self.menu_paper = RightHintMenu("용지 크기", mb)
        self.menu_view = RightHintMenu("보기", mb)
        self.menu_font = RightHintMenu("폰트", mb)
        self.menu_search = RightHintMenu("검색/표시", mb)
        self.menu_theme = RightHintMenu("용지색상", mb)

        mb.addMenu(self.menu_file)
        mb.addMenu(self.menu_paper)
        mb.addMenu(self.menu_view)
        mb.addMenu(self.menu_font)
        mb.addMenu(self.menu_search)
        mb.addMenu(self.menu_theme)

        mb.hovered.connect(self.on_menubar_hovered)
        mb.menuTitlePressed.connect(self.on_menu_title_pressed)

        for m in (self.menu_file, self.menu_paper, self.menu_view, self.menu_font, self.menu_search, self.menu_theme):
            m.aboutToShow.connect(self.on_any_menu_opened)

        act_open = QAction("불러오기", self); act_open.triggered.connect(self.open_file)
        act_folder = QAction("폴더", self); act_folder.triggered.connect(self.open_folder)
        self.menu_recent = RightHintMenu("최근", self.menu_file)
        self.refresh_recent_menu()

        act_exit = QAction("종료", self); act_exit.triggered.connect(self.close)

        # 인코딩으로 다시 열기 서브메뉴
        self.menu_encoding = RightHintMenu("인코딩으로 다시 열기 (TXT)", self.menu_file)
        for enc_name, enc_code in [
            ("자동 감지", ""),
            ("UTF-8", "utf-8"),
            ("CP949 / ANSI 한글 (Windows)", "cp949"),
            ("EUC-KR", "euc-kr"),
            ("UTF-16", "utf-16"),
            ("Latin-1 (서유럽)", "latin-1"),
        ]:
            if enc_name == "자동 감지":
                a = QAction(enc_name, self)
                a.triggered.connect(lambda: self.reload_with_encoding("") if hasattr(self, 'path') and self.path.lower().endswith('.txt') else None)
            else:
                a = QAction(enc_name, self)
                a.triggered.connect(lambda checked=False, e=enc_code: self.reload_with_encoding(e))
            self.menu_encoding.addAction(a)

        self.menu_file.addAction(act_open)
        self.menu_file.addAction(act_folder)
        self.menu_file.addMenu(self.menu_recent)

        self.menu_file.addSeparator()
        self.menu_file.addMenu(self.menu_encoding)
        self.menu_file.addSeparator()
        self.menu_file.addAction(act_exit)

        paper_group = QActionGroup(self); paper_group.setExclusive(True)
        self._paper_actions = []
        for name in PAPER_PRESETS.keys():
            a = QAction(name, self)
            a.setCheckable(True)
            a.setChecked(name == self.paper_name)
            a.triggered.connect(lambda checked=False, n=name: self.set_paper(n))
            paper_group.addAction(a)
            self.menu_paper.addAction(a)
            self._paper_actions.append(a)

        self.act_thumb = QAction("썸네일 표시", self)
        self.act_thumb.setCheckable(True)
        self.act_thumb.setChecked(self.thumb_visible)
        self.act_thumb.toggled.connect(self.toggle_thumbnails)
        self.menu_view.addAction(self.act_thumb)
        self.menu_view.addSeparator()

        view_group = QActionGroup(self); view_group.setExclusive(True)
        a1 = QAction("1장 보기", self); a1.setCheckable(True); a1.setChecked(self.view_mode == 1)
        a2 = QAction("2장 보기", self); a2.setCheckable(True); a2.setChecked(self.view_mode == 2)
        a1.triggered.connect(lambda: self.set_view_mode(1))
        a2.triggered.connect(lambda: self.set_view_mode(2))
        view_group.addAction(a1); view_group.addAction(a2)
        self.menu_view.addAction(a1); self.menu_view.addAction(a2)

        act_font = QAction("폰트 선택...", self); act_font.triggered.connect(self.open_font_dialog)
        act_bold = QAction("굵게(토글)", self); act_bold.setCheckable(True)
        act_bold.setChecked(self.text_layout.bold)
        act_bold.toggled.connect(self.set_text_bold)
        self.menu_font.addAction(act_font)
        self.menu_font.addAction(act_bold)

        act_find = QAction("검색...", self); act_find.setShortcut(QKeySequence.Find); act_find.triggered.connect(self.search_text)
        act_next = QAction("다음 찾기", self); act_next.setShortcut(QKeySequence("F3")); act_next.triggered.connect(self.find_next)
        act_prev = QAction("이전 찾기", self); act_prev.setShortcut(QKeySequence("Shift+F3")); act_prev.triggered.connect(self.find_prev)

        act_hl_add = QAction("선택 하이라이트 추가", self); act_hl_add.setShortcut(QKeySequence("Ctrl+H")); act_hl_add.triggered.connect(self.add_user_highlight)
        act_hl_clear = QAction("하이라이트 전체 삭제", self); act_hl_clear.triggered.connect(self.clear_user_highlights)

        act_bm_add = QAction("책갈피 추가", self); act_bm_add.setShortcut(QKeySequence("Ctrl+D")); act_bm_add.triggered.connect(self.add_bookmark)
        act_bm_list = QAction("책갈피 목록/이동", self); act_bm_list.setShortcut(QKeySequence("Ctrl+B")); act_bm_list.triggered.connect(self.goto_bookmark)
        act_bm_del = QAction("현재 책갈피 삭제", self); act_bm_del.triggered.connect(self.remove_bookmark_current)

        self.menu_search.addAction(act_find)
        self.menu_search.addAction(act_next)
        self.menu_search.addAction(act_prev)
        self.menu_search.addSeparator()
        self.menu_search.addAction(act_hl_add)
        self.menu_search.addAction(act_hl_clear)
        self.menu_search.addSeparator()
        self.menu_search.addAction(act_bm_add)
        self.menu_search.addAction(act_bm_list)
        self.menu_search.addAction(act_bm_del)

        theme_group = QActionGroup(self); theme_group.setExclusive(True)
        for tname in THEMES.keys():
            a = QAction(tname, self)
            a.setCheckable(True)
            a.setChecked(tname == self.theme_name)
            a.triggered.connect(lambda checked=False, n=tname: self.set_theme(n))
            theme_group.addAction(a)
            self.menu_theme.addAction(a)

    def build_ribbon(self):
        self.ribbon_stack = QStackedWidget()
        self.ribbon_stack.setFixedHeight(76)
        self.ribbon_stack.setStyleSheet("""
            QStackedWidget { background:#f5f5f5; border-bottom:1px solid #d0d0d0; }
            QWidget { color: #111111; }
            QComboBox { background: #ffffff; border: 1px solid #c0c0c0; padding: 2px 4px; }
            QComboBox QAbstractItemView { background: #ffffff; color: #111111; selection-background-color: #0078d7; selection-color: #ffffff; }
            QSpinBox, QDoubleSpinBox { background: #ffffff; border: 1px solid #c0c0c0; }
        """)

        # 0: 파일
        p0 = QWidget(); l0 = QHBoxLayout(p0); l0.setContentsMargins(8, 3, 8, 3); l0.setSpacing(10)
        b_open = RibbonButton("불러오기", QStyle.SP_DialogOpenButton, self); b_open.clicked.connect(self.open_file)
        b_folder = RibbonButton("폴더", QStyle.SP_DirOpenIcon, self); b_folder.clicked.connect(self.open_folder)
        b_recent = RibbonButton("최근", QStyle.SP_FileDialogDetailedView, self); b_recent.clicked.connect(self.open_recent_dialog)
        l0.addWidget(b_open); l0.addWidget(b_folder); l0.addWidget(b_recent)
        l0.addStretch()
        self.ribbon_stack.addWidget(p0)

        # 1: 용지 - 크기별 개별 리본 버튼
        p1 = QWidget(); l1 = QHBoxLayout(p1); l1.setContentsMargins(8, 3, 8, 3); l1.setSpacing(6)
        self._paper_btns: Dict[str, QToolButton] = {}

        # 용지별 라벨 단축명 + 아이콘 미니 사각형 비율 (w_mm, h_mm)
        _PAPER_META = {
            "전체보기(창맞춤)": ("전체맞춤", 1.0, 1.0),   # 정사각형 상징
            "A5 (소설)": ("A5",  148, 210),
            "B6 (문고)": ("B6",  128, 182),
            "Pocket (포켓)": ("Pocket", 105, 148),
            "Kindle 6\"": ("Kindle", 108, 144),
        }

        ICON_W, ICON_H = 28, 34   # 아이콘 커니스 최대 크기
        for pname in PAPER_PRESETS.keys():
            label, wv, hv = _PAPER_META.get(pname, (pname[:6], 148, 210))

            # 용지 비율에 맞춤 미니 사각형 커기 계산
            if isinstance(wv, float):          # 전체보기: 정사각형
                rw, rh = ICON_W - 4, ICON_H - 4
            else:
                ratio = float(wv) / float(hv)
                rh = ICON_H - 4
                rw = int(rh * ratio)
                if rw > ICON_W - 4:
                    rw = ICON_W - 4
                    rh = int(rw / ratio)

            pm = QPixmap(ICON_W, ICON_H)
            pm.fill(Qt.transparent)
            _p = QPainter(pm)
            _p.setRenderHints(QPainter.Antialiasing)
            ox = (ICON_W - rw) // 2
            oy = (ICON_H - rh) // 2
            # 용지 흰 배경
            _p.setBrush(QColor("#ffffff"))
            _p.setPen(QPen(QColor("#aaaaaa"), 1))
            _p.drawRect(ox, oy, rw, rh)
            # 텍스트 줄 패턴
            _p.setPen(QPen(QColor("#888888"), 1))
            line_gap = max(3, rh // 5)
            for i in range(1, 4):
                ly = oy + i * line_gap
                if ly < oy + rh - 2:
                    _p.drawLine(ox + 3, ly, ox + rw - 3, ly)
            _p.end()
            icon = QIcon(pm)

            btn = QToolButton()
            btn.setText(label)
            btn.setIcon(icon)
            btn.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
            btn.setAutoRaise(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setIconSize(QSize(ICON_W, ICON_H))
            btn.setFixedSize(72, 64)
            btn.setCheckable(True)
            btn.setChecked(pname == self.paper_name)
            btn.setToolTip(pname)
            btn.setStyleSheet("""
                QToolButton {
                    border: 1px solid transparent;
                    border-radius: 8px;
                    padding: 2px;
                }
                QToolButton:hover {
                    background-color: #deeaf7;
                    border: 1px solid #0078d7;
                }
                QToolButton:checked {
                    background-color: #cce4ff;
                    border: 2px solid #0078d7;
                }
            """)
            btn.clicked.connect(lambda checked=False, n=pname: self.set_paper(n))
            l1.addWidget(btn)
            self._paper_btns[pname] = btn

        l1.addStretch()
        self.ribbon_stack.addWidget(p1)

        # 2: 보기
        p2 = QWidget(); l2 = QHBoxLayout(p2); l2.setContentsMargins(8, 3, 8, 3); l2.setSpacing(10)
        b1 = RibbonButton("1장", QStyle.SP_FileIcon, self); b1.clicked.connect(lambda: self.set_view_mode(1))
        b2 = RibbonButton("2장", QStyle.SP_FileLinkIcon, self); b2.clicked.connect(lambda: self.set_view_mode(2))
        self.thumb_btn = RibbonButton("썸네일", QStyle.SP_ComputerIcon, self)
        self.thumb_btn.setCheckable(True)
        self.thumb_btn.setChecked(self.thumb_visible)
        self.thumb_btn.clicked.connect(lambda: self.toggle_thumbnails(self.thumb_btn.isChecked()))
        l2.addWidget(b1); l2.addWidget(b2); l2.addWidget(self.thumb_btn)
        l2.addStretch()
        self.ribbon_stack.addWidget(p2)

        # 3: 폰트/레이아웃
        p3 = QWidget(); l3 = QHBoxLayout(p3); l3.setContentsMargins(8, 3, 8, 3); l3.setSpacing(10)
        b_font = RibbonButton("폰트", QStyle.SP_DialogYesButton, self); b_font.clicked.connect(self.open_font_dialog)
        self.bold_btn = RibbonButton("굵게", QStyle.SP_TitleBarShadeButton, self)
        self.bold_btn.setCheckable(True)
        self.bold_btn.setChecked(self.text_layout.bold)
        self.bold_btn.clicked.connect(lambda: self.set_text_bold(self.bold_btn.isChecked()))
        l3.addWidget(b_font); l3.addWidget(self.bold_btn)

        b_minus = RibbonButton("A-", QStyle.SP_ArrowDown, self)
        b_plus = RibbonButton("A+", QStyle.SP_ArrowUp, self)
        self.sp_font_size = QSpinBox()
        self.sp_font_size.setRange(6, 72)
        self.sp_font_size.setValue(int(self.text_layout.font_size))
        self.sp_font_size.setFixedWidth(70)
        self.sp_font_size.valueChanged.connect(self.set_text_font_size)
        b_minus.clicked.connect(lambda: self.set_text_font_size(int(self.text_layout.font_size) - 1))
        b_plus.clicked.connect(lambda: self.set_text_font_size(int(self.text_layout.font_size) + 1))

        l3.addWidget(QLabel("크기"))
        l3.addWidget(b_minus); l3.addWidget(self.sp_font_size); l3.addWidget(b_plus)

        self.sp_ui_font = QSpinBox()
        self.sp_ui_font.setRange(8, 30)
        self.sp_ui_font.setValue(int(self.settings.get("ui_font_size", 12)))
        self.sp_ui_font.setFixedWidth(50)
        self.sp_ui_font.valueChanged.connect(self.set_ui_font_size)
        l3.addWidget(QLabel(" UI크기")); l3.addWidget(self.sp_ui_font)

        self.sp_margin_lr = QSpinBox(); self.sp_margin_lr.setRange(0, 200); self.sp_margin_lr.setValue(self.text_layout.margin_lr)
        self.sp_margin_lr.valueChanged.connect(self.set_margin_lr); self.sp_margin_lr.setFixedWidth(70)
        self.sp_margin_tb = QSpinBox(); self.sp_margin_tb.setRange(0, 200); self.sp_margin_tb.setValue(self.text_layout.margin_tb)
        self.sp_margin_tb.valueChanged.connect(self.set_margin_tb); self.sp_margin_tb.setFixedWidth(70)
        self.sp_line = QDoubleSpinBox(); self.sp_line.setRange(1.0, 2.5); self.sp_line.setSingleStep(0.05); self.sp_line.setValue(self.text_layout.line_spacing)
        self.sp_line.valueChanged.connect(self.set_line_spacing); self.sp_line.setFixedWidth(70)
        self.sp_para = QSpinBox(); self.sp_para.setRange(0, 60); self.sp_para.setValue(self.text_layout.para_spacing)
        self.sp_para.valueChanged.connect(self.set_para_spacing); self.sp_para.setFixedWidth(70)
        self.sp_indent = QDoubleSpinBox(); self.sp_indent.setRange(0.0, 9999.0); self.sp_indent.setDecimals(2); self.sp_indent.setSingleStep(0.25); self.sp_indent.setValue(float(self.text_layout.indent))
        self.sp_indent.valueChanged.connect(self.set_para_indent); self.sp_indent.setFixedWidth(70)

        l3.addWidget(QLabel("좌우")); l3.addWidget(self.sp_margin_lr)
        l3.addWidget(QLabel("상하")); l3.addWidget(self.sp_margin_tb)
        l3.addWidget(QLabel("줄")); l3.addWidget(self.sp_line)
        l3.addWidget(QLabel("문단")); l3.addWidget(self.sp_para)
        l3.addWidget(QLabel("들여쓰기")); l3.addWidget(self.sp_indent)
        # SpinBox/Label 높이 상한 고정: 폰트 변경 시 팽창 방지
        SPIN_H = 32
        for sp in (self.sp_font_size, self.sp_ui_font,
                   self.sp_margin_lr, self.sp_margin_tb,
                   self.sp_line, self.sp_para, self.sp_indent):
            sp.setMaximumHeight(SPIN_H)
        for lbl in p3.findChildren(QLabel):
            lbl.setMaximumHeight(SPIN_H)

        l3.addStretch()
        self.ribbon_stack.addWidget(p3)

        # 4: 검색/표시
        p4 = QWidget(); l4 = QHBoxLayout(p4); l4.setContentsMargins(8, 3, 8, 3); l4.setSpacing(10)
        b_find = RibbonButton("검색", QStyle.SP_FileDialogContentsView, self); b_find.clicked.connect(self.search_text)
        b_next = RibbonButton("다음", QStyle.SP_ArrowForward, self); b_next.clicked.connect(self.find_next)
        b_prev = RibbonButton("이전", QStyle.SP_ArrowBack, self); b_prev.clicked.connect(self.find_prev)
        b_hl = RibbonButton("형광펜", QStyle.SP_DialogApplyButton, self); b_hl.clicked.connect(self.add_user_highlight)
        b_bm = RibbonButton("책갈피", QStyle.SP_DirHomeIcon, self); b_bm.clicked.connect(self.add_bookmark)
        l4.addWidget(b_find); l4.addWidget(b_prev); l4.addWidget(b_next); l4.addWidget(b_hl); l4.addWidget(b_bm)
        l4.addStretch()
        self.ribbon_stack.addWidget(p4)

        # 5: 용지색상 - 테마별 개별 리본 버튼
        p5 = QWidget(); l5 = QHBoxLayout(p5); l5.setContentsMargins(8, 3, 8, 3); l5.setSpacing(8)
        self._theme_btns: Dict[str, QToolButton] = {}

        # 테마이름 → (아이콘 배경색, 아이콘 텍스트 색, 테두리 색) 매핑
        _THEME_META = {
            "화이트": ("#ffffff", "#555555", "#bbbbbb"),
            "블랙": ("#111111", "#eeeeee", "#444444"),
            "종이":  ("#fdf6e3", "#6b5b2e", "#c8b98a"),
        }

        for tname, tcolor in THEMES.items():
            bg, fg, border = _THEME_META.get(tname, (tcolor, "#333333", "#999999"))

            # 색상 스와치 아이콘 생성 (28×28 콜러 사각형)
            pm = QPixmap(28, 28)
            pm.fill(QColor(bg))
            _p = QPainter(pm)
            _p.setRenderHints(QPainter.Antialiasing)
            _p.setPen(QPen(QColor(border), 1.5))
            _p.setBrush(Qt.NoBrush)
            _p.drawRoundedRect(1, 1, 25, 25, 4, 4)
            # 간단한 줄 패턴으로 텍스트를 시각화
            _p.setPen(QPen(QColor(fg), 1.5))
            for _y in (9, 14, 19):
                _p.drawLine(5, _y, 22, _y)
            _p.end()
            icon = QIcon(pm)

            btn = QToolButton()
            btn.setText(tname)
            btn.setIcon(icon)
            btn.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
            btn.setAutoRaise(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setIconSize(QSize(28, 28))
            btn.setFixedSize(68, 58)
            btn.setCheckable(True)
            btn.setChecked(tname == self.theme_name)
            btn.setStyleSheet("""
                QToolButton {
                    border: 1px solid transparent;
                    border-radius: 8px;
                    padding: 2px;
                }
                QToolButton:hover {
                    background-color: #deeaf7;
                    border: 1px solid #0078d7;
                }
                QToolButton:checked {
                    background-color: #cce4ff;
                    border: 2px solid #0078d7;
                }
            """)
            btn.clicked.connect(lambda checked=False, n=tname: self.set_theme(n))
            l5.addWidget(btn)
            self._theme_btns[tname] = btn

        l5.addStretch()
        self.ribbon_stack.addWidget(p5)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    f = safe_ui_font()
    f.setPointSize(12)
    QApplication.setFont(f)

    w = EPUBReader()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
