import sys, os, json, time, hashlib, html
from datetime import datetime
from typing import Callable

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTextEdit, QWidget, QLabel, QToolBar,
    QComboBox, QSpinBox, QColorDialog, QStatusBar, QFrame, QMenu, QPushButton,
    QHBoxLayout, QVBoxLayout, QListWidget, QListWidgetItem, QMenuBar, QLineEdit, QStyle
)
from PyQt6.QtGui import (
    QAction, QActionGroup, QTextCharFormat, QColor, QFont, QPainter,
    QLinearGradient, QBrush, QTextCursor, QFontMetrics, QTextDocument,
    QTextOption, QFontDatabase, QPixmap, QIcon, QPolygon
)
from PyQt6.QtCore import Qt, QTimer, QRect, QPoint, QSize

from ebooklib import epub, ITEM_DOCUMENT
from bs4 import BeautifulSoup


# ------------------------------------------------------------
# Menu bar: click only, no hover auto-open
# ------------------------------------------------------------
class HoverMenuBar(QMenuBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.menu_hover_callback: Callable[[str], None] | None = None
        self._current_action: QAction | None = None
        self._opened_menu: QMenu | None = None

    def _restore_style_after_click(self):
        parent = self.window()
        callback = getattr(parent, "apply_ui_font_size", None)

        if callable(callback):
            callback()

    def mouseMoveEvent(self, a0):
        if a0 is None:
            return
        super().mouseMoveEvent(a0)

    def mousePressEvent(self, a0):
        if a0 is None:
            return
        action = self.actionAt(a0.pos())

        menu = action.menu() if action is not None else None
        if action is not None and menu is not None:

            if self.menu_hover_callback:
                self.menu_hover_callback(action.text())

            if self._opened_menu and self._opened_menu is not menu:
                self._opened_menu.close()

            self._current_action = action
            self._opened_menu = menu
            self.setActiveAction(action)

            rect = self.actionGeometry(action)
            pos = self.mapToGlobal(rect.bottomLeft())

            try:
                menu.aboutToHide.disconnect(self._on_menu_hidden)
            except Exception:
                pass

            menu.aboutToHide.connect(self._on_menu_hidden)
            menu.popup(pos)

            QTimer.singleShot(0, self._restore_style_after_click)

            a0.accept()
            return

        super().mousePressEvent(a0)

    def _on_menu_hidden(self):
        self._current_action = None
        self._opened_menu = None
        self.setActiveAction(None)


# ------------------------------------------------------------
# Page widgets
# ------------------------------------------------------------
class PageTextEdit(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.next_page_callback: Callable[[], None] | None = None
        self.prev_page_callback: Callable[[], None] | None = None
        self.context_menu_callback: Callable[[QPoint], None] | None = None

        self.setReadOnly(True)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.setWordWrapMode(QTextOption.WrapMode.WrapAtWordBoundaryOrAnywhere)

    def wheelEvent(self, e):
        if e is None:
            return
        if e.angleDelta().y() < 0 and self.next_page_callback:
            self.next_page_callback()
        elif e.angleDelta().y() > 0 and self.prev_page_callback:
            self.prev_page_callback()
        e.accept()

    def keyPressEvent(self, e):
        if e is None:
            return
        key = e.key()
        if key in [Qt.Key.Key_Right, Qt.Key.Key_Down, Qt.Key.Key_PageDown, Qt.Key.Key_Space]:
            if self.next_page_callback:
                self.next_page_callback()
            return
        if key in [Qt.Key.Key_Left, Qt.Key.Key_Up, Qt.Key.Key_PageUp, Qt.Key.Key_Backspace]:
            if self.prev_page_callback:
                self.prev_page_callback()
            return

        parent = self.window()
        toggle = getattr(parent, "toggle_fullscreen", None)
        if key == Qt.Key.Key_F11 and callable(toggle):
            toggle()
            return
        if parent is not None and key == Qt.Key.Key_Escape and parent.isFullScreen():
            parent.showNormal()
            return

        super().keyPressEvent(e)

    def contextMenuEvent(self, e):
        if e is None:
            return
        if self.context_menu_callback:
            self.context_menu_callback(e.globalPos())
            e.accept()
            return
        super().contextMenuEvent(e)


class PageThumbnailList(QListWidget):
    """Thumbnail list wheel/key navigation linked to page turning."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.next_page_callback: Callable[[], None] | None = None
        self.prev_page_callback: Callable[[], None] | None = None

    def wheelEvent(self, e):
        if e is None:
            return
        delta = e.angleDelta().y()
        if delta < 0 and self.next_page_callback:
            self.next_page_callback()
        elif delta > 0 and self.prev_page_callback:
            self.prev_page_callback()
        e.accept()

    def keyPressEvent(self, e):
        if e is None:
            return
        key = e.key()
        if key in [Qt.Key.Key_Right, Qt.Key.Key_Down, Qt.Key.Key_PageDown, Qt.Key.Key_Space]:
            if self.next_page_callback:
                self.next_page_callback()
            return
        if key in [Qt.Key.Key_Left, Qt.Key.Key_Up, Qt.Key.Key_PageUp, Qt.Key.Key_Backspace]:
            if self.prev_page_callback:
                self.prev_page_callback()
            return
        super().keyPressEvent(e)


class PdfPageLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.next_page_callback: Callable[[], None] | None = None
        self.prev_page_callback: Callable[[], None] | None = None
        self.context_menu_callback: Callable[[QPoint], None] | None = None
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet("background-color:#202020; border:none;")

    def wheelEvent(self, a0):
        if a0 is None:
            return
        if a0.angleDelta().y() < 0 and self.next_page_callback:
            self.next_page_callback()
        elif a0.angleDelta().y() > 0 and self.prev_page_callback:
            self.prev_page_callback()
        a0.accept()

    def keyPressEvent(self, ev):
        if ev is None:
            return
        key = ev.key()
        if key in [Qt.Key.Key_Right, Qt.Key.Key_Down, Qt.Key.Key_PageDown, Qt.Key.Key_Space]:
            if self.next_page_callback:
                self.next_page_callback()
            return
        if key in [Qt.Key.Key_Left, Qt.Key.Key_Up, Qt.Key.Key_PageUp, Qt.Key.Key_Backspace]:
            if self.prev_page_callback:
                self.prev_page_callback()
            return
        super().keyPressEvent(ev)

    def contextMenuEvent(self, ev):
        if ev is None:
            return
        if self.context_menu_callback:
            self.context_menu_callback(ev.globalPos())
            ev.accept()
            return
        super().contextMenuEvent(ev)


# ------------------------------------------------------------
# Canvas
# ------------------------------------------------------------
class BookCanvas(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.view_mode = 1
        self.paper_color = QColor("#f4ecd8")
        self.text_color = QColor("#2b2118")
        self.outer_color = QColor("#3a3028")
        self.paper_rects = []
        self.page_numbers = []
        self.edge_click_callback: Callable[[str], None] | None = None
        self.context_menu_callback: Callable[[QPoint], None] | None = None
        self.show_bookmark = False
        self.bookmark_side = "right"
        self.pdf_mode = False

    def set_view_mode(self, mode):
        self.view_mode = mode
        self.update()

    def set_pdf_mode(self, enabled):
        self.pdf_mode = enabled
        self.update()

    def set_theme(self, paper, text, outer):
        self.paper_color = QColor(paper)
        self.text_color = QColor(text)
        self.outer_color = QColor(outer)
        self.update()

    def set_paper_rects(self, rects):
        self.paper_rects = rects
        self.update()

    def set_page_numbers(self, numbers):
        self.page_numbers = numbers
        self.update()

    def set_bookmark(self, visible, side="right"):
        self.show_bookmark = visible
        self.bookmark_side = side
        self.update()

    def contextMenuEvent(self, a0):
        if a0 is None:
            return
        if self.context_menu_callback:
            self.context_menu_callback(a0.globalPos())
            a0.accept()
            return
        super().contextMenuEvent(a0)

    def mousePressEvent(self, a0):
        if a0 is None:
            return
        if a0.button() == Qt.MouseButton.RightButton:
            return super().mousePressEvent(a0)

        if not self.edge_click_callback:
            return

        x = a0.position().x()
        edge = max(70, int(self.width() * 0.07))

        if x <= edge:
            self.edge_click_callback("prev")
        elif x >= self.width() - edge:
            self.edge_click_callback("next")

    def paintEvent(self, a0):
        if a0 is None:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        if self.pdf_mode:
            painter.fillRect(self.rect(), QColor("#202020"))
            return

        painter.fillRect(self.rect(), self.outer_color)

        if not self.paper_rects:
            return

        if self.view_mode == 1:
            self.draw_single_page(painter)
        else:
            self.draw_spread_pages(painter)

        self.draw_page_numbers(painter)

    def draw_single_page(self, painter):
        rect = self.paper_rects[0]
        painter.fillRect(rect, self.paper_color)

        shadow = QLinearGradient(rect.right(), 0, rect.left(), 0)
        shadow.setColorAt(0.0, QColor(0, 0, 0, 65))
        shadow.setColorAt(1.0, QColor(0, 0, 0, 0))
        painter.fillRect(rect, shadow)

        painter.setPen(QColor(120, 100, 80))
        painter.drawRect(rect)

        self.draw_bookmark(painter, rect)
        self.draw_page_edges(painter, rect, "single")

    def draw_spread_pages(self, painter):
        if len(self.paper_rects) < 2:
            return

        left_rect, right_rect = self.paper_rects
        painter.fillRect(left_rect, self.paper_color)
        painter.fillRect(right_rect, self.paper_color)

        center = left_rect.right()
        center_shadow_width = 28
        center_shadow = QLinearGradient(center - center_shadow_width, 0, center + center_shadow_width, 0)
        center_shadow.setColorAt(0.0, QColor(0, 0, 0, 0))
        center_shadow.setColorAt(0.45, QColor(0, 0, 0, 55))
        center_shadow.setColorAt(0.55, QColor(0, 0, 0, 55))
        center_shadow.setColorAt(1.0, QColor(0, 0, 0, 0))
        painter.fillRect(
            QRect(center - center_shadow_width, left_rect.top(), center_shadow_width * 2, left_rect.height()),
            center_shadow
        )

        # left_shadow = QLinearGradient(left_rect.right(), 0, left_rect.left(), 0)
        # left_shadow.setColorAt(0.0, QColor(0, 0, 0, 65))
        # left_shadow.setColorAt(1.0, QColor(0, 0, 0, 0))
        # painter.fillRect(left_rect, left_shadow)

        # right_shadow = QLinearGradient(right_rect.left(), 0, right_rect.right(), 0)
        # right_shadow.setColorAt(0.0, QColor(0, 0, 0, 65))
        # right_shadow.setColorAt(1.0, QColor(0, 0, 0, 0))
        # painter.fillRect(right_rect, right_shadow)

        painter.setPen(QColor(120, 100, 80))
        painter.drawRect(left_rect)
        painter.drawRect(right_rect)

        target_rect = left_rect if self.bookmark_side == "left" else right_rect
        self.draw_bookmark(painter, target_rect)

        self.draw_page_edges(painter, left_rect, "left")
        self.draw_page_edges(painter, right_rect, "right")

    def draw_bookmark(self, painter, page_rect):
        if not self.show_bookmark:
            return

        ribbon_width = 26
        ribbon_height = 110
        x = page_rect.left() + 30 if self.bookmark_side == "left" else page_rect.right() - 56
        y = page_rect.top()

        painter.setBrush(QColor("#b91c1c"))
        painter.setPen(Qt.PenStyle.NoPen)
        points = [
            QPoint(x, y), QPoint(x + ribbon_width, y),
            QPoint(x + ribbon_width, y + ribbon_height),
            QPoint(x + ribbon_width // 2, y + ribbon_height - 20),
            QPoint(x, y + ribbon_height),
        ]
        painter.drawPolygon(QPolygon(points))

    def draw_page_edges(self, painter, rect, side="single"):
        painter.setPen(QColor(180, 160, 130, 120))

        if side in ["single", "right"]:
            for i in range(1, 9):
                painter.drawLine(rect.right() - i, rect.top() + 8, rect.right() - i, rect.bottom() - 8)

        if side in ["single", "left"]:
            for i in range(1, 9):
                painter.drawLine(rect.left() + i, rect.top() + 8, rect.left() + i, rect.bottom() - 8)

        painter.setPen(QColor(210, 190, 160, 90))
        for i in range(1, 9):
            painter.drawLine(rect.left() + 8, rect.top() + i, rect.right() - 8, rect.top() + i)
            painter.drawLine(rect.left() + 8, rect.bottom() - i, rect.right() - 8, rect.bottom() - i)

    def draw_page_numbers(self, painter):
        if not self.page_numbers or not self.paper_rects:
            return

        painter.save()
        painter.setFont(QFont("맑은 고딕", 10))
        painter.setPen(QColor(self.text_color.red(), self.text_color.green(), self.text_color.blue(), 180))

        for info in self.page_numbers:
            i = info.get("rect_index", 0)
            if i >= len(self.paper_rects):
                continue

            rect = self.paper_rects[i]
            number_rect = QRect(rect.left() + 24, rect.bottom() - 34, rect.width() - 48, 24)
            align = info.get("align", "center")

            flags = Qt.AlignmentFlag.AlignVCenter
            if align == "left":
                flags |= Qt.AlignmentFlag.AlignLeft
            elif align == "right":
                flags |= Qt.AlignmentFlag.AlignRight
            else:
                flags |= Qt.AlignmentFlag.AlignCenter

            painter.drawText(number_rect, flags, info.get("text", ""))

        painter.restore()


# ------------------------------------------------------------
# Main Window
# ------------------------------------------------------------
class DocumentReader(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python Document Reader (2026.05.14)")
        self.resize(1280, 840)

        self.current_file = ""
        self.book_id = ""
        self.full_text = ""
        self.pages = []
        self.current_page = 0
        self.is_pdf = False
        self.pdf_doc = None
        self.pdf_page_count = 0

        self.view_mode = 1
        self.paper_size = "자동"
        self.fit_to_window = True
        self.inner_margin_x = 24
        self.inner_margin_y = 24
        self.paragraph_spacing = 8
        self.first_line_indent = 18
        self.line_height = 1.55
        self.font_bold = False
        self.ui_font_size = 10
        self.page_thumbnail_visible = True
        self.paper_theme = "종이 색상"
        self.app_font_name = "맑은 고딕"
        self.app_font_size = 14

        self.bookmark_page = None
        self.bookmark_side = "right"
        self.annotations = []
        self.highlight_color = QColor("yellow")
        self.underline_color = QColor("#1d4ed8")

        self.start_time = time.time()
        self.saved_read_seconds = 0

        self.recent_books = []
        self.max_recent_books = 10

        self.search_keyword = ""
        self.search_results = []
        self.current_search_index = -1

        self.base_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
        self.data_dir = os.path.join(self.base_dir, "reader_data")
        os.makedirs(self.data_dir, exist_ok=True)
        self.app_config_file = os.path.join(self.data_dir, "app_config.json")
        self.last_open_dir = os.path.expanduser("~")
        self.loading_state = False

        self.load_app_config()
        self.init_ui()
        self.init_timer()

    # ---------------- UI ----------------
    def init_ui(self):
        self.create_menu()
        self.create_context_menu()
        self.create_hover_menu_bar()
        self.create_toolbar()

        self.canvas = BookCanvas()
        self.canvas.edge_click_callback = self.handle_edge_click
        self.canvas.context_menu_callback = self.show_edit_context_menu
        self.viewer_container = self.canvas

        self.left_viewer = PageTextEdit(self.viewer_container)
        self.right_viewer = PageTextEdit(self.viewer_container)
        self.left_pdf_label = PdfPageLabel(self.viewer_container)
        self.right_pdf_label = PdfPageLabel(self.viewer_container)

        for w in [self.left_viewer, self.right_viewer, self.left_pdf_label, self.right_pdf_label]:
            w.next_page_callback = self.next_page
            w.prev_page_callback = self.prev_page
            w.context_menu_callback = self.show_edit_context_menu
            w.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        for viewer in [self.left_viewer, self.right_viewer]:
            font = QFont(self.app_font_name, self.app_font_size)
            font.setBold(self.font_bold)
            viewer.setFont(font)

        self.main_layout_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_layout_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        self.create_page_thumbnail_panel()

        self.content_widget = QWidget()
        self.content_layout = QHBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(0)
        self.content_layout.addWidget(self.page_thumb_panel)
        self.content_layout.addWidget(self.viewer_container, 1)

        self.setMenuBar(self.hover_menu_bar)
        self.main_layout.addWidget(self.toolbar)
        self.main_layout.addWidget(self.content_widget, 1)

        self.setCentralWidget(self.main_layout_widget)
        self.create_bottom_log()

        self.apply_global_ui_config_to_widgets()
        self.apply_ui_font_size()
        self.apply_text_style()
        self.update_recent_books_menu()
        self.update_page_geometry()
        self.display_page()

        QTimer.singleShot(0, self.force_initial_two_page_layout)
        QTimer.singleShot(100, self.force_initial_two_page_layout)

    def create_menu(self):
        self.file_menu = QMenu("파일", self)
        self.edit_menu = QMenu("편집", self)
        self.view_menu = QMenu("보기", self)
        self.paper_menu = QMenu("용지 크기", self)
        self.margin_menu = QMenu("여백", self)
        self.paragraph_menu = QMenu("문단", self)
        self.color_menu = QMenu("용지 색상", self)
        self.bookmark_menu = QMenu("책갈피", self)
        self.recent_menu = QMenu("최근 열어본 문서", self)

        self.file_menu.addAction(self.action("파일 불러오기", self.open_document))
        self.file_menu.addAction(self.action("폴더 불러오기", self.open_library_folder))
        self.file_menu.addMenu(self.recent_menu)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.action("저장", self.save_state))
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.action("종료", self.close))

        self.edit_menu.addAction(self.action("색깔 표시", self.highlight_text))
        self.edit_menu.addAction(self.action("밑줄", self.underline_text))
        self.edit_menu.addSeparator()
        self.edit_menu.addAction(self.action("색상 선택", self.choose_highlight_color))
        self.edit_menu.addAction(self.action("밑줄 색상 선택", self.choose_underline_color))
        self.edit_menu.addSeparator()
        self.edit_menu.addAction(self.action("현재 페이지 표시 삭제", self.clear_current_page_annotations))

        self.view_menu.addAction(self.action("1장 보기", lambda: self.set_view_mode_from_menu(1)))
        self.view_menu.addAction(self.action("2장 보기", lambda: self.set_view_mode_from_menu(2)))
        self.view_menu.addSeparator()

        self.thumb_panel_action = QAction("페이지 썸네일 표시", self)
        self.thumb_panel_action.setCheckable(True)
        self.thumb_panel_action.setChecked(self.page_thumbnail_visible)
        self.thumb_panel_action.triggered.connect(self.toggle_page_thumbnail_panel)
        self.view_menu.addAction(self.thumb_panel_action)

        self.bold_menu_action = QAction("본문 굵게", self)
        self.bold_menu_action.setCheckable(True)
        self.bold_menu_action.setChecked(self.font_bold)
        self.bold_menu_action.triggered.connect(self.toggle_font_bold)
        self.view_menu.addAction(self.bold_menu_action)

        self.view_menu.addAction(self.action("전체 화면 전환", self.toggle_fullscreen))

        ui_font_menu = QMenu("메뉴/로그 글자 크기", self.view_menu)
        self.view_menu.addMenu(ui_font_menu)
        self.ui_font_action_group = QActionGroup(self)
        self.ui_font_action_group.setExclusive(True)
        for size in [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 30]:
            a = QAction(f"{size} pt", self)
            a.setCheckable(True)
            a.setChecked(size == self.ui_font_size)
            a.triggered.connect(lambda checked=False, s=size: self.set_ui_font_size_from_menu(s))
            self.ui_font_action_group.addAction(a)
            ui_font_menu.addAction(a)

        self.paper_size_action_group = QActionGroup(self)
        self.paper_size_action_group.setExclusive(True)
        for name in ["자동", "국판", "신국판", "문고판", "A5", "A4", "태블릿", "넓게"]:
            a = QAction(name, self)
            a.setCheckable(True)
            a.setChecked(name == self.paper_size)
            a.triggered.connect(lambda checked=False, n=name: self.change_paper_size(n))
            self.paper_size_action_group.addAction(a)
            self.paper_menu.addAction(a)

        self.fit_to_window_action = QAction("창에 맞게 확대", self)
        self.fit_to_window_action.setCheckable(True)
        self.fit_to_window_action.setChecked(self.fit_to_window)
        self.fit_to_window_action.triggered.connect(self.toggle_fit_to_window)
        self.paper_menu.addSeparator()
        self.paper_menu.addAction(self.fit_to_window_action)

        x_menu = QMenu("좌우 여백", self.margin_menu)
        self.margin_menu.addMenu(x_menu)
        y_menu = QMenu("상하 여백", self.margin_menu)
        self.margin_menu.addMenu(y_menu)
        for value in [8, 12, 16, 24, 32, 40, 56, 72]:
            x_menu.addAction(self.action(f"{value}px", lambda checked=False, v=value: self.change_inner_margin_x(v)))
            y_menu.addAction(self.action(f"{value}px", lambda checked=False, v=value: self.change_inner_margin_y(v)))

        spacing_menu = QMenu("문단 간격", self.paragraph_menu)
        self.paragraph_menu.addMenu(spacing_menu)
        indent_menu = QMenu("첫 줄 들여쓰기", self.paragraph_menu)
        self.paragraph_menu.addMenu(indent_menu)
        line_menu = QMenu("줄 간격", self.paragraph_menu)
        self.paragraph_menu.addMenu(line_menu)
        for value in [0, 4, 8, 12, 16, 20, 24, 32]:
            spacing_menu.addAction(self.action(f"{value}px", lambda checked=False, v=value: self.change_paragraph_spacing(v)))
        for value in [0, 8, 12, 18, 24, 32, 40, 56]:
            indent_menu.addAction(self.action(f"{value}px", lambda checked=False, v=value: self.change_first_line_indent(v)))
        for text, value in [("1.2", 1.2), ("1.35", 1.35), ("1.55", 1.55), ("1.75", 1.75), ("2.0", 2.0)]:
            line_menu.addAction(self.action(text, lambda checked=False, v=value: self.change_line_height(v)))

        for name in ["화이트", "블랙", "종이 색상"]:
            self.color_menu.addAction(self.action(name, lambda checked=False, n=name: self.change_paper_color(n)))

        self.bookmark_menu.addAction(self.action("책갈피 지정", self.set_bookmark))
        self.bookmark_menu.addAction(self.action("왼쪽 페이지 책갈피", self.set_bookmark_left))
        self.bookmark_menu.addAction(self.action("오른쪽 페이지 책갈피", self.set_bookmark_right))
        self.bookmark_menu.addAction(self.action("책갈피 이동", self.go_bookmark))
        self.bookmark_menu.addSeparator()
        self.bookmark_menu.addAction(self.action("책갈피 해제", self.clear_bookmark))

    def create_context_menu(self):
        self.context_menu = QMenu(self)
        self.context_menu.addAction(self.action("색깔 표시", self.highlight_text))
        self.context_menu.addAction(self.action("밑줄", self.underline_text))
        self.context_menu.addSeparator()
        self.context_menu.addAction(self.action("색상 선택", self.choose_highlight_color))
        self.context_menu.addAction(self.action("밑줄 색상 선택", self.choose_underline_color))
        self.context_menu.addSeparator()
        self.context_menu.addAction(self.action("현재 페이지 표시 삭제", self.clear_current_page_annotations))
        self.context_menu.addSeparator()

        bm = QMenu("책갈피", self.context_menu)
        self.context_menu.addMenu(bm)
        bm.addAction(self.action("현재 페이지 책갈피", self.set_bookmark))
        bm.addAction(self.action("왼쪽 페이지 책갈피", self.set_bookmark_left))
        bm.addAction(self.action("오른쪽 페이지 책갈피", self.set_bookmark_right))
        bm.addAction(self.action("책갈피 이동", self.go_bookmark))
        bm.addSeparator()
        bm.addAction(self.action("책갈피 해제", self.clear_bookmark))

    def create_hover_menu_bar(self):
        self.hover_menu_bar = HoverMenuBar(self)
        self.hover_menu_bar.setNativeMenuBar(False)  # 추가
        self.hover_menu_bar.setObjectName("HoverMenuBar")

        for m in [
            self.file_menu, self.edit_menu, self.view_menu, self.paper_menu,
            self.margin_menu, self.paragraph_menu, self.color_menu, self.bookmark_menu
        ]:
            self.hover_menu_bar.addMenu(m)
        self.hover_menu_bar.menu_hover_callback = self.change_ribbon_by_menu

    def create_toolbar(self):
        self.toolbar = QToolBar("리본 메뉴")
        self.toolbar.setMovable(False)
        self.toolbar.setFloatable(False)
        self.toolbar.setIconSize(QSize(22, 22))
        self.toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.toolbar.setFixedHeight(78)
        self.ribbon_widgets = []
        self.build_file_ribbon()

    def action(self, text, slot, icon=None):
        a = QAction(icon or QIcon(), text, self)
        a.triggered.connect(slot)
        return a

    def clear_ribbon(self):
        self.toolbar.clear()
        for w in getattr(self, "ribbon_widgets", []):
            try:
                w.deleteLater()
            except Exception:
                pass
        self.ribbon_widgets = []

    def rw(self, widget):
        self.ribbon_widgets.append(widget)
        self.toolbar.addWidget(widget)
        return widget

    def standard_icon(self, icon: QStyle.StandardPixmap) -> QIcon:
        style = self.style()
        if style is None:
            return QIcon()
        return style.standardIcon(icon)

    def ra(self, icon, text, slot):
        a = QAction(self.standard_icon(icon), text, self)
        a.triggered.connect(slot)
        self.toolbar.addAction(a)
        return a

    def rl(self, text):
        return self.rw(QLabel(text))

    def change_ribbon_by_menu(self, name):
        mapping = {
            "파일": self.build_file_ribbon,
            "편집": self.build_edit_ribbon,
            "보기": self.build_view_ribbon,
            "용지 크기": self.build_paper_ribbon,
            "여백": self.build_margin_ribbon,
            "문단": self.build_paragraph_ribbon,
            "용지 색상": self.build_color_ribbon,
            "책갈피": self.build_bookmark_ribbon,
        }
        if name in mapping:
            mapping[name]()

    def build_file_ribbon(self):
        self.clear_ribbon()
        S = QStyle.StandardPixmap
        self.ra(S.SP_DialogOpenButton, "열기", self.open_document)
        self.ra(S.SP_DirIcon, "폴더", self.open_library_folder)
        self.ra(S.SP_DialogSaveButton, "저장", self.save_state)
        self.toolbar.addSeparator()

        self.rl("최근 문서 ")
        combo = QComboBox()
        combo.setMinimumWidth(280)
        combo.addItem("최근 문서 선택", "")
        for p in [p for p in self.recent_books if os.path.exists(p)]:
            combo.addItem(os.path.basename(p), p)
        combo.currentIndexChanged.connect(lambda i, c=combo: self.open_recent_book(c.currentData()) if c.currentData() else None)
        self.rw(combo)

        self.toolbar.addSeparator()
        self.ra(S.SP_DialogCloseButton, "종료", self.close)
        self.apply_ui_font_size()

    def build_edit_ribbon(self):
        self.clear_ribbon()
        S = QStyle.StandardPixmap
        self.ra(S.SP_DialogApplyButton, "표시", self.highlight_text)
        self.ra(S.SP_FileDialogDetailedView, "밑줄", self.underline_text)
        self.ra(S.SP_DialogHelpButton, "표시색", self.choose_highlight_color)
        self.ra(S.SP_DialogHelpButton, "밑줄색", self.choose_underline_color)
        self.ra(S.SP_TrashIcon, "표시삭제", self.clear_current_page_annotations)
        self.toolbar.addSeparator()

        self.rl("검색 ")
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("검색어 입력")
        self.search_edit.setMinimumWidth(220)
        self.search_edit.returnPressed.connect(self.search_next)
        self.rw(self.search_edit)
        self.ra(S.SP_ArrowUp, "이전검색", self.search_prev)
        self.ra(S.SP_ArrowDown, "다음검색", self.search_next)
        self.apply_ui_font_size()

    def build_view_ribbon(self):
        self.clear_ribbon()
        S = QStyle.StandardPixmap
        self.ra(S.SP_ArrowBack, "이전", self.prev_page)
        self.ra(S.SP_ArrowForward, "다음", self.next_page)
        self.ra(S.SP_TitleBarMaxButton, "전체화면", self.toggle_fullscreen)
        self.toolbar.addSeparator()

        self.rl("화면 ")
        self.view_combo = QComboBox()
        self.view_combo.addItems(["1장 보기", "2장 보기"])
        self.view_combo.setCurrentIndex(0 if self.view_mode == 1 else 1)
        self.view_combo.currentIndexChanged.connect(self.change_view_mode)
        self.rw(self.view_combo)

        self.bold_action = QAction(self.standard_icon(S.SP_DialogApplyButton), "굵게", self)
        self.bold_action.setCheckable(True)
        self.bold_action.setChecked(self.font_bold)
        self.bold_action.triggered.connect(self.toggle_font_bold)
        self.toolbar.addAction(self.bold_action)

        self.thumb_ribbon_action = QAction(self.standard_icon(S.SP_FileDialogListView), "페이지 썸네일", self)
        self.thumb_ribbon_action.setCheckable(True)
        self.thumb_ribbon_action.setChecked(self.page_thumbnail_visible)
        self.thumb_ribbon_action.triggered.connect(self.toggle_page_thumbnail_panel)
        self.toolbar.addAction(self.thumb_ribbon_action)

        self.apply_ui_font_size()

    def build_paper_ribbon(self):
        self.clear_ribbon()
        S = QStyle.StandardPixmap
        self.rl("용지 ")
        combo = QComboBox()
        combo.addItems(["자동", "국판", "신국판", "문고판", "A5", "A4", "태블릿", "넓게"])
        combo.setCurrentText(self.paper_size)
        combo.currentTextChanged.connect(self.change_paper_size)
        self.rw(combo)

        fit_action = QAction(self.standard_icon(S.SP_DesktopIcon), "창맞춤", self)
        fit_action.setCheckable(True)
        fit_action.setChecked(self.fit_to_window)
        fit_action.triggered.connect(self.toggle_fit_to_window)
        self.toolbar.addAction(fit_action)
        self.apply_ui_font_size()

    def build_margin_ribbon(self):
        self.clear_ribbon()
        for label, current, slot in [
            ("좌우 여백 ", self.inner_margin_x, self.change_inner_margin_x),
            ("상하 여백 ", self.inner_margin_y, self.change_inner_margin_y),
        ]:
            self.rl(label)
            combo = QComboBox()
            for value in [8, 12, 16, 24, 32, 40, 56, 72]:
                combo.addItem(f"{value}px", value)
            idx = combo.findData(current)
            combo.setCurrentIndex(max(idx, 0))
            combo.currentIndexChanged.connect(lambda i, c=combo, s=slot: s(c.currentData()))
            self.rw(combo)
            self.toolbar.addSeparator()
        self.apply_ui_font_size()

    def build_paragraph_ribbon(self):
        self.clear_ribbon()
        self.rl("폰트 ")
        self.font_combo = QComboBox()
        self.font_combo.setMinimumWidth(180)
        self.font_combo.addItems(self.get_system_fonts())
        if self.font_combo.findText(self.app_font_name) >= 0:
            self.font_combo.setCurrentText(self.app_font_name)
        self.font_combo.currentTextChanged.connect(self.change_font)
        self.rw(self.font_combo)

        self.rl(" 크기 ")
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 60)
        self.font_size_spin.setValue(self.app_font_size)
        self.font_size_spin.valueChanged.connect(self.change_font_size)
        self.rw(self.font_size_spin)

        for label, current, values, slot in [
            (" 문단 ", self.paragraph_spacing, [0, 4, 8, 12, 16, 20, 24, 32], self.change_paragraph_spacing),
            (" 들여쓰기 ", self.first_line_indent, [0, 8, 12, 18, 24, 32, 40, 56], self.change_first_line_indent),
            (" 줄간격 ", self.line_height, [1.2, 1.35, 1.55, 1.75, 2.0], self.change_line_height),
        ]:
            self.rl(label)
            combo = QComboBox()
            for value in values:
                combo.addItem(str(value), value)
            idx = combo.findData(current)
            combo.setCurrentIndex(max(idx, 0))
            combo.currentIndexChanged.connect(lambda i, c=combo, s=slot: s(c.currentData()))
            self.rw(combo)

        self.apply_ui_font_size()

    def build_color_ribbon(self):
        self.clear_ribbon()
        S = QStyle.StandardPixmap
        for name in ["화이트", "블랙", "종이 색상"]:
            self.ra(S.SP_DriveFDIcon, name, lambda checked=False, n=name: self.change_paper_color(n))
        self.apply_ui_font_size()

    def build_bookmark_ribbon(self):
        self.clear_ribbon()
        S = QStyle.StandardPixmap
        self.ra(S.SP_BrowserReload, "책갈피 지정", self.set_bookmark)
        self.ra(S.SP_ArrowBack, "왼쪽 책갈피", self.set_bookmark_left)
        self.ra(S.SP_ArrowForward, "오른쪽 책갈피", self.set_bookmark_right)
        self.ra(S.SP_ArrowUp, "책갈피 이동", self.go_bookmark)
        self.ra(S.SP_TrashIcon, "책갈피 해제", self.clear_bookmark)
        self.apply_ui_font_size()

    def create_page_thumbnail_panel(self):
        self.page_thumb_panel = QWidget()
        self.page_thumb_panel.setObjectName("PageThumbnailPanel")
        layout = QVBoxLayout(self.page_thumb_panel)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        header = QHBoxLayout()
        title = QLabel("페이지 썸네일")
        close_btn = QPushButton("×")
        close_btn.setFixedWidth(28)
        close_btn.clicked.connect(lambda: self.toggle_page_thumbnail_panel(False))
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(close_btn)

        self.page_thumb_list = PageThumbnailList()
        self.page_thumb_list.setMinimumWidth(200)
        self.page_thumb_list.setMaximumWidth(260)
        self.page_thumb_list.setIconSize(QSize(150, 210))
        self.page_thumb_list.setSpacing(10)
        self.page_thumb_list.next_page_callback = self.next_page
        self.page_thumb_list.prev_page_callback = self.prev_page
        self.page_thumb_list.itemClicked.connect(self.go_page_from_thumbnail)

        layout.addLayout(header)
        layout.addWidget(self.page_thumb_list, 1)
        self.page_thumb_panel.setVisible(self.page_thumbnail_visible)

    def create_bottom_log(self):
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.file_label = QLabel("파일: 없음")
        self.page_label = QLabel("페이지: 0 / 0")
        self.progress_label = QLabel("진행율: 0%")
        self.time_label = QLabel("현재 시간: -")
        self.read_time_label = QLabel("읽은 시간: 00:00:00")
        self.status.addWidget(self.file_label, 3)
        self.status.addPermanentWidget(self.page_label)
        self.status.addPermanentWidget(self.progress_label)
        self.status.addPermanentWidget(self.time_label)
        self.status.addPermanentWidget(self.read_time_label)

    def init_timer(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_time_log)
        self.timer.start(1000)

    # ---------------- Document loading ----------------
    def open_document(self):
        start_dir = self.last_open_dir if os.path.exists(self.last_open_dir) else os.path.expanduser("~")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "문서 파일 열기",
            start_dir,
            "Documents (*.epub *.txt *.pdf);;EPUB Files (*.epub);;Text Files (*.txt);;PDF Files (*.pdf)"
        )
        if file_path:
            self.open_document_file(file_path)

    def open_library_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "문서 폴더 선택", self.last_open_dir)
        if not folder:
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "폴더 내 문서 선택",
            folder,
            "Documents (*.epub *.txt *.pdf);;EPUB Files (*.epub);;Text Files (*.txt);;PDF Files (*.pdf)"
        )
        if file_path:
            self.open_document_file(file_path)

    def open_document_file(self, file_path):
        self.last_open_dir = os.path.dirname(file_path)
        self.current_file = file_path
        self.book_id = hashlib.md5(os.path.abspath(file_path).encode("utf-8")).hexdigest()
        self.add_recent_book(file_path)

        self.current_page = 0
        self.bookmark_page = None
        self.bookmark_side = "right"
        self.annotations = []
        self.saved_read_seconds = 0
        self.start_time = time.time()

        ext = os.path.splitext(file_path)[1].lower()
        self.is_pdf = ext == ".pdf"
        self.canvas.set_pdf_mode(self.is_pdf)

        if self.is_pdf:
            self.open_pdf(file_path)
            self.full_text = ""
            self.pages = []
        else:
            self.close_pdf()
            self.full_text = self.load_document_text(file_path)

        self.file_label.setText("파일: " + os.path.basename(file_path))

        QTimer.singleShot(0, self.finish_open_document)
        QTimer.singleShot(120, self.finish_open_document)

    def finish_open_document(self):
        self.apply_reader_font()
        self.update_page_geometry()

        if not self.is_pdf:
            self.repaginate()

        self.load_state()

        if self.view_mode == 2 and self.current_page % 2 == 1:
            self.current_page -= 1
        self.current_page = max(0, self.current_page)

        self.update_page_geometry()
        self.display_page()
        self.update_page_thumbnails()

        if self.view_mode == 2:
            self.force_initial_two_page_layout()

        self.save_app_config()

    def load_document_text(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".epub":
            return self.load_epub_text(file_path)
        if ext == ".txt":
            return self.load_txt_text(file_path)
        return ""

    def load_epub_text(self, file_path):
        book = epub.read_epub(file_path)
        item_map = {item.get_id(): item for item in book.get_items()}
        text = ""

        for spine_item in book.spine:
            item = item_map.get(spine_item[0])
            if item and item.get_type() == ITEM_DOCUMENT:
                soup = BeautifulSoup(item.get_content(), "html.parser")
                text += soup.get_text("\n") + "\n\n"

        if not text.strip():
            for item in book.get_items():
                if item.get_type() == ITEM_DOCUMENT:
                    soup = BeautifulSoup(item.get_content(), "html.parser")
                    text += soup.get_text("\n") + "\n\n"

        return text.strip()

    def load_txt_text(self, file_path):
        for enc in ["utf-8", "utf-8-sig", "cp949", "euc-kr"]:
            try:
                with open(file_path, "r", encoding=enc) as f:
                    return f.read().strip()
            except UnicodeDecodeError:
                continue
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read().strip()

    def open_pdf(self, file_path):
        try:
            import pymupdf as fitz
        except ImportError:
            self.status.showMessage("PDF 보기를 위해 pymupdf 설치가 필요합니다.", 5000)
            self.is_pdf = False
            return

        self.close_pdf()
        self.pdf_doc = fitz.open(file_path)
        self.pdf_page_count = len(self.pdf_doc)

    def close_pdf(self):
        try:
            if self.pdf_doc:
                self.pdf_doc.close()
        except Exception:
            pass
        self.pdf_doc = None
        self.pdf_page_count = 0

    # ---------------- Geometry ----------------
    def paper_ratio(self):
        if self.paper_size == "국판":
            return 148 / 210
        if self.paper_size == "신국판":
            return 152 / 225
        if self.paper_size == "문고판":
            return 105 / 148
        if self.paper_size == "A5":
            return 148 / 210
        if self.paper_size == "A4":
            return 210 / 297
        if self.paper_size == "태블릿":
            return 3 / 4
        if self.paper_size == "넓게":
            return 4 / 3
        return None

    def paper_scale(self):
        if self.fit_to_window:
            return 1.0
        if self.paper_size == "문고판":
            return 0.72
        if self.paper_size in ["국판", "A5"]:
            return 0.82
        if self.paper_size == "신국판":
            return 0.86
        if self.paper_size in ["A4", "태블릿"]:
            return 0.90
        return 1.0

    def calculate_page_rects(self):
        w = max(self.canvas.width(), 400)
        h = max(self.canvas.height(), 300)
        outer_margin = 28
        gap = 16 if self.is_pdf else 4

        available_w = max(w - outer_margin * 2, 300)
        available_h = max(h - outer_margin * 2, 300)

        if self.is_pdf:
            if self.view_mode == 1:
                return [QRect(outer_margin, outer_margin, available_w, available_h)]
            page_w = max(220, (available_w - gap) // 2)
            return [
                QRect(outer_margin, outer_margin, page_w, available_h),
                QRect(outer_margin + page_w + gap, outer_margin, page_w, available_h)
            ]

        ratio = self.paper_ratio()
        if ratio is None:
            ratio = 148 / 210
        scale = self.paper_scale()

        if self.view_mode == 1:
            if self.paper_size == "자동":
                page_w = int(available_w * 0.96)
                page_h = int(available_h * 0.96)
            else:
                max_w = int(available_w * scale)
                max_h = int(available_h * scale)
                page_w = min(max_w, int(max_h * ratio))
                page_h = int(page_w / ratio)
                if page_h > max_h:
                    page_h = max_h
                    page_w = int(page_h * ratio)

            page_w = max(260, min(page_w, available_w))
            page_h = max(320, min(page_h, available_h))
            return [QRect((w - page_w) // 2, (h - page_h) // 2, page_w, page_h)]

        max_page_w = int((available_w - gap) / 2)
        if self.paper_size == "자동":
            page_w = max_page_w
            page_h = int(available_h * 0.96)
        else:
            max_page_h = int(available_h * scale)
            page_w = min(int(max_page_w * scale), int(max_page_h * ratio))
            page_h = int(page_w / ratio)
            if page_h > max_page_h:
                page_h = max_page_h
                page_w = int(page_h * ratio)

        page_w = max(220, min(page_w, max_page_w))
        page_h = max(300, min(page_h, available_h))
        total_w = page_w * 2 + gap
        x0 = (w - total_w) // 2
        y = (h - page_h) // 2
        return [QRect(x0, y, page_w, page_h), QRect(x0 + page_w + gap, y, page_w, page_h)]

    def update_page_geometry(self):
        if not hasattr(self, "canvas"):
            return

        rects = self.calculate_page_rects()
        self.canvas.set_paper_rects(rects)
        self.canvas.set_pdf_mode(self.is_pdf)

        text_widgets = [self.left_viewer, self.right_viewer]
        pdf_widgets = [self.left_pdf_label, self.right_pdf_label]

        if self.is_pdf:
            for w in text_widgets:
                w.hide()
            if self.view_mode == 1:
                rect = rects[0]
                self.left_pdf_label.hide()
                self.right_pdf_label.show()
                self.right_pdf_label.setGeometry(rect)
                self.right_pdf_label.setFixedSize(rect.width(), rect.height())
            else:
                left_rect, right_rect = rects
                self.left_pdf_label.show()
                self.right_pdf_label.show()
                self.left_pdf_label.setGeometry(left_rect)
                self.right_pdf_label.setGeometry(right_rect)
                self.left_pdf_label.setFixedSize(left_rect.width(), left_rect.height())
                self.right_pdf_label.setFixedSize(right_rect.width(), right_rect.height())
                self.left_pdf_label.raise_()
                self.right_pdf_label.raise_()
        else:
            for w in pdf_widgets:
                w.hide()
            if self.view_mode == 1:
                rect = rects[0]
                self.left_viewer.hide()
                self.right_viewer.show()
                self.right_viewer.setGeometry(rect)
                self.right_viewer.setFixedSize(rect.width(), rect.height())
            else:
                left_rect, right_rect = rects
                self.left_viewer.show()
                self.right_viewer.show()
                self.left_viewer.setGeometry(left_rect)
                self.right_viewer.setGeometry(right_rect)
                self.left_viewer.setFixedSize(left_rect.width(), left_rect.height())
                self.right_viewer.setFixedSize(right_rect.width(), right_rect.height())
                self.left_viewer.raise_()
                self.right_viewer.raise_()

        self.apply_text_style()

    def force_initial_two_page_layout(self):
        if self.view_mode != 2:
            return

        self.update_page_geometry()

        if self.is_pdf:
            self.left_pdf_label.show()
            self.right_pdf_label.show()
            self.left_pdf_label.raise_()
            self.right_pdf_label.raise_()
        else:
            self.left_viewer.show()
            self.right_viewer.show()
            if not self.pages:
                self.left_viewer.setPlainText("문서를 불러와 주세요.")
                self.right_viewer.setPlainText("")
            self.left_viewer.raise_()
            self.right_viewer.raise_()

        self.canvas.set_view_mode(2)
        self.canvas.update()

    # ---------------- Text pagination/display ----------------
    def is_cjk_char(self, ch):
        code = ord(ch)
        return (
            0xAC00 <= code <= 0xD7A3 or
            0x1100 <= code <= 0x11FF or
            0x3130 <= code <= 0x318F or
            0x4E00 <= code <= 0x9FFF or
            0x3040 <= code <= 0x309F or
            0x30A0 <= code <= 0x30FF
        )

    def add_cjk_break_opportunities(self, text):
        result = []
        for i, ch in enumerate(text):
            result.append(ch)
            if i < len(text) - 1 and self.is_cjk_char(ch) and self.is_cjk_char(text[i + 1]):
                result.append("\u200b")
        return "".join(result)

    def text_to_html(self, text):
        paragraphs = [p.strip() for p in text.splitlines() if p.strip()]
        body = ""
        for p in paragraphs:
            p = self.add_cjk_break_opportunities(p)
            escaped = html.escape(p)
            body += (
                f'<p style="margin-top:0px; margin-bottom:{self.paragraph_spacing}px; '
                f'text-indent:{self.first_line_indent}px; line-height:{self.line_height}; '
                f'text-align:justify; text-justify:inter-ideograph; white-space:normal;">'
                f'{escaped}</p>'
            )

        return f"""
        <html>
        <head>
        <style>
            body {{ margin:0; padding:0; }}
            p {{
                text-align: justify;
                text-justify: inter-ideograph;
                white-space: normal;
            }}
        </style>
        </head>
        <body>{body}</body>
        </html>
        """

    def text_fits_in_page(self, text, page_width, page_height, font):
        doc = QTextDocument()
        doc.setDefaultFont(font)
        doc.setHtml(self.text_to_html(text))
        doc.setDocumentMargin(0)

        option = doc.defaultTextOption()
        option.setAlignment(Qt.AlignmentFlag.AlignJustify)
        option.setWrapMode(QTextOption.WrapMode.WrapAtWordBoundaryOrAnywhere)
        doc.setDefaultTextOption(option)

        available_width = max(page_width - self.inner_margin_x * 2, 80)
        available_height = max(page_height - self.inner_margin_y * 2, 100)
        doc.setTextWidth(available_width)
        return doc.size().height() <= available_height * 0.96

    def find_page_end_by_layout(self, text, start, page_width, page_height, font):
        remaining = len(text) - start
        if remaining <= 0:
            return start

        fm = QFontMetrics(font)
        line_height = max(fm.lineSpacing(), 1)
        avg_char_width = max(fm.averageCharWidth(), 7)

        available_width = max(page_width - self.inner_margin_x * 2, 80)
        available_height = max(page_height - self.inner_margin_y * 2, 100)

        chars_per_line = max(int(available_width / avg_char_width), 8)
        lines_per_page = max(int(available_height / line_height), 5)
        rough_chars = max(int(chars_per_line * lines_per_page * 0.75), 80)

        low = 1
        high = min(remaining, rough_chars * 2)

        while high < remaining:
            candidate = text[start:start + high].strip()
            if not candidate:
                break
            if self.text_fits_in_page(candidate, page_width, page_height, font):
                high = min(remaining, high * 2)
            else:
                break

        high = min(high, remaining)
        best = low

        while low <= high:
            mid = (low + high) // 2
            candidate = text[start:start + mid].strip()
            if self.text_fits_in_page(candidate, page_width, page_height, font):
                best = mid
                low = mid + 1
            else:
                high = mid - 1

        end = start + best
        if end < len(text):
            search_start = start + int(best * 0.55)
            cut = max(
                text.rfind("\n", search_start, end),
                text.rfind(".", search_start, end),
                text.rfind("。", search_start, end),
                text.rfind("!", search_start, end),
                text.rfind("?", search_start, end),
                text.rfind(" ", search_start, end),
            )
            if cut > start:
                end = cut + 1

        if end <= start:
            end = min(start + max(best, 1), len(text))

        return end

    def repaginate(self):
        if self.is_pdf:
            return

        if not self.full_text:
            self.pages = []
            return

        old_progress = 0
        if self.pages:
            old_progress = self.current_page / max(len(self.pages) - 1, 1)

        self.update_page_geometry()
        viewer = self.right_viewer if self.view_mode == 1 else self.left_viewer

        page_width = max(viewer.width(), 260)
        page_height = max(viewer.height(), 320)
        font = viewer.font()

        text = self.full_text
        pages = []
        start = 0

        while start < len(text):
            while start < len(text) and text[start].isspace():
                start += 1
            if start >= len(text):
                break

            end = self.find_page_end_by_layout(text, start, page_width, page_height, font)
            page_text = text[start:end].strip()
            if page_text:
                pages.append(page_text)
            start = end if end > start else start + 1

        self.pages = pages if pages else [""]
        self.current_page = min(int(old_progress * max(len(self.pages) - 1, 1)), max(len(self.pages) - 1, 0))

        if self.view_mode == 2 and self.current_page % 2 == 1:
            self.current_page -= 1

        self.search_results = []
        self.current_search_index = -1

    def display_page(self):
        self.update_page_geometry()

        if self.is_pdf:
            self.display_pdf_page()
            return

        self.left_viewer.clear()
        self.right_viewer.clear()

        if not self.pages:
            if self.view_mode == 1:
                self.left_viewer.hide()
                self.right_viewer.show()
                self.right_viewer.setPlainText("문서를 불러와 주세요.")
            else:
                self.left_viewer.show()
                self.right_viewer.show()
                self.left_viewer.setPlainText("문서를 불러와 주세요.")
                self.right_viewer.setPlainText("")
                self.left_viewer.raise_()
                self.right_viewer.raise_()

            self.update_view_mode()
            self.update_canvas_page_numbers()
            self.update_log()
            return

        self.current_page = max(0, min(self.current_page, len(self.pages) - 1))

        if self.view_mode == 2:
            if self.current_page % 2 == 1:
                self.current_page -= 1

            self.left_viewer.show()
            self.right_viewer.show()

            left_page = self.current_page
            right_page = self.current_page + 1

            self.left_viewer.setHtml(self.text_to_html(self.pages[left_page]))
            self.apply_annotations_to_viewer(self.left_viewer, left_page)

            if right_page < len(self.pages):
                self.right_viewer.setHtml(self.text_to_html(self.pages[right_page]))
                self.apply_annotations_to_viewer(self.right_viewer, right_page)
            else:
                self.right_viewer.setHtml("")

            self.left_viewer.raise_()
            self.right_viewer.raise_()
        else:
            self.left_viewer.hide()
            self.right_viewer.show()
            self.right_viewer.setHtml(self.text_to_html(self.pages[self.current_page]))
            self.apply_annotations_to_viewer(self.right_viewer, self.current_page)

        self.reset_viewer_scroll(self.left_viewer)
        self.reset_viewer_scroll(self.right_viewer)
        self.update_view_mode()
        self.update_canvas_page_numbers()
        self.update_log()
        self.update_current_thumbnail_selection()

    def reset_viewer_scroll(self, viewer):
        cursor = viewer.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        viewer.setTextCursor(cursor)
        vertical = viewer.verticalScrollBar()
        if vertical is not None:
            vertical.setValue(0)
        horizontal = viewer.horizontalScrollBar()
        if horizontal is not None:
            horizontal.setValue(0)
        viewer.ensureCursorVisible()
        vertical = viewer.verticalScrollBar()
        if vertical is not None:
            vertical.setValue(0)

    # ---------------- PDF display ----------------
    def render_pdf_page_pixmap(self, page_index, target_size):
        if not self.pdf_doc or page_index < 0 or page_index >= self.pdf_page_count:
            return QPixmap()

        page = self.pdf_doc[page_index]
        rect = page.rect
        if rect.width <= 0 or rect.height <= 0:
            return QPixmap()

        tw, th = max(target_size.width(), 50), max(target_size.height(), 50)
        scale = min(tw / rect.width, th / rect.height) if self.fit_to_window else 1.0
        scale = max(scale, 0.1)

        matrix = None
        try:
            import pymupdf as fitz
            matrix = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image_format = QPixmap()
            qimg = QPixmap()
            # Use QImage conversion
            from PyQt6.QtGui import QImage
            fmt = QImage.Format.Format_RGB888
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()
            return QPixmap.fromImage(img)
        except Exception as e:
            print("PDF 렌더링 오류:", e)
            return QPixmap()

    def display_pdf_page(self):
        self.current_page = max(0, min(self.current_page, max(self.pdf_page_count - 1, 0)))

        if self.view_mode == 2 and self.current_page % 2 == 1:
            self.current_page -= 1

        self.left_pdf_label.clear()
        self.right_pdf_label.clear()

        if not self.pdf_doc or self.pdf_page_count == 0:
            if self.view_mode == 1:
                self.left_pdf_label.hide()
                self.right_pdf_label.show()
                self.right_pdf_label.setText("PDF 문서를 불러와 주세요.")
            else:
                self.left_pdf_label.show()
                self.right_pdf_label.show()
                self.left_pdf_label.setText("PDF 문서를 불러와 주세요.")
                self.right_pdf_label.setText("")
            self.update_log()
            return

        if self.view_mode == 1:
            self.left_pdf_label.hide()
            self.right_pdf_label.show()
            pix = self.render_pdf_page_pixmap(self.current_page, self.right_pdf_label.size())
            self.right_pdf_label.setPixmap(pix)
        else:
            self.left_pdf_label.show()
            self.right_pdf_label.show()
            left_page = self.current_page
            right_page = self.current_page + 1

            left_pix = self.render_pdf_page_pixmap(left_page, self.left_pdf_label.size())
            self.left_pdf_label.setPixmap(left_pix)

            if right_page < self.pdf_page_count:
                right_pix = self.render_pdf_page_pixmap(right_page, self.right_pdf_label.size())
                self.right_pdf_label.setPixmap(right_pix)
            else:
                self.right_pdf_label.setPixmap(QPixmap())

            self.left_pdf_label.raise_()
            self.right_pdf_label.raise_()

        self.update_view_mode()
        self.update_canvas_page_numbers()
        self.update_log()
        self.update_current_thumbnail_selection()

    # ---------------- Style ----------------
    def apply_reader_font(self):
        for viewer in [self.left_viewer, self.right_viewer]:
            font = QFont(self.app_font_name, self.app_font_size)
            font.setBold(self.font_bold)
            viewer.setFont(font)

    def apply_text_style(self):
        if not hasattr(self, "left_viewer"):
            return

        text_color = self.canvas.text_color.name()
        style = f"""
            QTextEdit {{
                background-color: transparent;
                color: {text_color};
                padding: 0px;
                border: none;
            }}
        """

        for viewer in [self.left_viewer, self.right_viewer]:
            viewer.setStyleSheet(style)

            doc = viewer.document()
            if doc is None:
                continue

            doc.setDocumentMargin(0)

            option = doc.defaultTextOption()
            option.setAlignment(Qt.AlignmentFlag.AlignJustify)
            option.setWrapMode(
                QTextOption.WrapMode.WrapAtWordBoundaryOrAnywhere
            )
            doc.setDefaultTextOption(option)

            viewer.setAlignment(Qt.AlignmentFlag.AlignJustify)
            viewer.setViewportMargins(
                self.inner_margin_x,
                self.inner_margin_y,
                self.inner_margin_x,
                self.inner_margin_y
            )
            viewer.setVerticalScrollBarPolicy(
                Qt.ScrollBarPolicy.ScrollBarAlwaysOff
            )
            viewer.setHorizontalScrollBarPolicy(
                Qt.ScrollBarPolicy.ScrollBarAlwaysOff
            )

        for label in [self.left_pdf_label, self.right_pdf_label]:
            label.setStyleSheet("background-color:#202020; border:none; color:white;")

    def apply_ui_font_size(self):
        font = QFont("맑은 고딕", self.ui_font_size)
        self.setFont(font)

        menu_bar_height = max(34, int(self.ui_font_size * 2.8))
        menu_item_min_width = max(72, int(self.ui_font_size * 7.2))
        menu_item_padding_v = max(7, int(self.ui_font_size * 0.65))
        menu_item_padding_h = max(18, int(self.ui_font_size * 1.8))

        menu_bar_style = f"""
            QMenuBar#HoverMenuBar {{
                background-color: #f0f0f0;
                border-bottom: 1px solid #c8c8c8;
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
            }}

            QMenuBar#HoverMenuBar::item {{
                background: transparent;
                padding: 5px 10px;
                color: #111111;
            }}

            QMenuBar#HoverMenuBar::item:selected {{
                background-color: #dcdcdc;
            }}

            QMenuBar#HoverMenuBar::item:pressed {{
                background-color: #cfcfcf;
            }}
        """

        toolbar_style = f"""
            QToolBar {{
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
                background-color: #fafafa;
                border-bottom: 1px solid #d0d0d0;
                spacing: 6px;
                padding: 4px;
                min-height: 70px;
                max-height: 78px;
            }}
            QToolButton {{
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
                padding: 4px 8px;
                min-width: 54px;
            }}
            QToolButton:hover {{
                background-color: #e5e5e5;
            }}
            QLabel, QComboBox, QSpinBox, QLineEdit {{
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
            }}
            QComboBox {{
                min-width: 120px;
            }}
            QSpinBox {{
                min-width: 60px;
            }}
        """

        menu_style = f"""
            QMenu {{
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
            }}
            QMenu::item {{
                padding: 6px 32px 6px 24px;
            }}
            QMenu::item:selected {{
                background: #dcdcdc;
            }}
        """

        thumb_style = f"""
            QWidget#PageThumbnailPanel {{
                background-color: #f7f7f7;
                border-right: 1px solid #cfcfcf;
            }}
            QWidget#PageThumbnailPanel QLabel {{
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
                font-weight: bold;
            }}
            QWidget#PageThumbnailPanel QPushButton {{
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
                border: 1px solid #cccccc;
                background-color: #ffffff;
            }}
            QListWidget {{
                font-family: "맑은 고딕";
                font-size: {self.ui_font_size}pt;
                background-color: #eeeeee;
                border: 1px solid #d0d0d0;
                outline: none;
            }}
            QListWidget::item {{
                padding: 8px;
                margin: 4px;
                border-radius: 4px;
            }}
            QListWidget::item:selected {{
                background-color: #c7ddff;
                border: 1px solid #4f8df5;
            }}
        """

        if hasattr(self, "hover_menu_bar"):
            self.hover_menu_bar.setFont(QFont("맑은 고딕", self.ui_font_size))
            self.hover_menu_bar.setFixedHeight(menu_bar_height)
            self.hover_menu_bar.setStyleSheet(menu_bar_style)

        if hasattr(self, "toolbar"):
            self.toolbar.setStyleSheet(toolbar_style)
            self.toolbar.setIconSize(QSize(22, 22))
            self.toolbar.setFixedHeight(78)
            for widget in self.toolbar.findChildren(QWidget):
                widget.setFont(font)

        for menu in [
            getattr(self, "file_menu", None), getattr(self, "edit_menu", None),
            getattr(self, "view_menu", None), getattr(self, "paper_menu", None),
            getattr(self, "margin_menu", None), getattr(self, "paragraph_menu", None),
            getattr(self, "color_menu", None), getattr(self, "bookmark_menu", None),
            getattr(self, "recent_menu", None), getattr(self, "context_menu", None),
        ]:
            if menu:
                menu.setStyleSheet(menu_style)
                menu.setFont(font)

        if hasattr(self, "page_thumb_panel"):
            self.page_thumb_panel.setStyleSheet(thumb_style)

        if hasattr(self, "status"):
            self.status.setFont(font)
            for widget in self.status.findChildren(QWidget):
                widget.setFont(font)

    # ---------------- Thumbnails ----------------
    def create_page_thumbnail_pixmap(self, page_text, page_number):
        thumb_w, thumb_h = 150, 210
        pixmap = QPixmap(thumb_w, thumb_h)
        pixmap.fill(Qt.GlobalColor.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        shadow_rect = QRect(8, 8, thumb_w - 12, thumb_h - 12)
        painter.fillRect(shadow_rect, QColor(0, 0, 0, 35))

        page_rect = QRect(4, 4, thumb_w - 14, thumb_h - 14)
        painter.fillRect(page_rect, self.canvas.paper_color)
        painter.setPen(QColor(170, 150, 120))
        painter.drawRect(page_rect)

        painter.setPen(self.canvas.text_color)
        title_font = QFont("맑은 고딕", 8)
        title_font.setBold(True)
        painter.setFont(title_font)
        painter.drawText(
            QRect(page_rect.left() + 8, page_rect.top() + 6, page_rect.width() - 16, 18),
            Qt.AlignmentFlag.AlignCenter,
            f"{page_number}쪽"
        )

        body_font = QFont(self.app_font_name, 7)
        body_font.setBold(self.font_bold)
        painter.setFont(body_font)

        text_rect = QRect(page_rect.left() + 12, page_rect.top() + 30, page_rect.width() - 24, page_rect.height() - 44)
        preview = page_text.replace("\n", " ").strip()[:260]
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop | Qt.TextFlag.TextWordWrap, preview)

        painter.end()
        return pixmap

    def create_pdf_thumbnail_pixmap(self, page_index):
        if not self.pdf_doc:
            return QPixmap(150, 210)
        pix = self.render_pdf_page_pixmap(page_index, QSize(150, 210))
        canvas = QPixmap(150, 210)
        canvas.fill(Qt.GlobalColor.transparent)
        p = QPainter(canvas)
        p.fillRect(QRect(8, 8, 138, 198), QColor(0, 0, 0, 45))
        if not pix.isNull():
            scaled = pix.scaled(136, 196, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            x = (150 - scaled.width()) // 2
            y = (210 - scaled.height()) // 2
            p.drawPixmap(x, y, scaled)
        p.end()
        return canvas

    def update_page_thumbnails(self):
        if not hasattr(self, "page_thumb_list"):
            return

        self.page_thumb_list.blockSignals(True)
        self.page_thumb_list.clear()
        self.page_thumb_list.setIconSize(QSize(150, 210))
        self.page_thumb_list.setSpacing(10)

        total = self.pdf_page_count if self.is_pdf else len(self.pages)
        for index in range(total):
            pixmap = self.create_pdf_thumbnail_pixmap(index) if self.is_pdf else self.create_page_thumbnail_pixmap(self.pages[index], index + 1)
            item = QListWidgetItem()
            item.setIcon(QIcon(pixmap))
            item.setText(f"{index + 1}쪽")
            item.setData(Qt.ItemDataRole.UserRole, index)
            item.setSizeHint(QSize(180, 245))
            self.page_thumb_list.addItem(item)

        self.page_thumb_list.blockSignals(False)
        self.update_current_thumbnail_selection()

    def update_current_thumbnail_selection(self):
        total = self.pdf_page_count if self.is_pdf else len(self.pages)
        if not hasattr(self, "page_thumb_list") or total == 0:
            return
        row = max(0, min(self.current_page, self.page_thumb_list.count() - 1))
        self.page_thumb_list.blockSignals(True)
        self.page_thumb_list.setCurrentRow(row)
        item = self.page_thumb_list.item(row)
        if item:
            self.page_thumb_list.scrollToItem(item, QListWidget.ScrollHint.PositionAtCenter)
        self.page_thumb_list.blockSignals(False)

    def go_page_from_thumbnail(self, item):
        page_index = item.data(Qt.ItemDataRole.UserRole)
        if page_index is None:
            return
        self.current_page = int(page_index)
        if self.view_mode == 2 and self.current_page % 2 == 1:
            self.current_page -= 1
        self.display_page()
        self.save_state()

    # ---------------- Search ----------------
    def normalize_for_search(self, s):
        return (
            s.replace("\u200b", "")
             .replace("\u2029", "\n")
             .replace("\xa0", " ")
             .casefold()
        )

    def make_normalized_index_map(self, text):
        normalized = []
        index_map = []

        for i, ch in enumerate(text):
            if ch == "\u200b":
                continue
            if ch == "\u2029":
                ch = "\n"
            elif ch == "\xa0":
                ch = " "

            folded = ch.casefold()
            for folded_ch in folded:
                normalized.append(folded_ch)
                index_map.append(i)

        return "".join(normalized), index_map

    def build_search_results(self, keyword):
        self.search_results = []
        self.current_search_index = -1

        if self.is_pdf:
            self.status.showMessage("PDF 검색은 현재 지원하지 않습니다.", 3000)
            return

        if not keyword or not self.pages:
            return

        keyword_normal = self.normalize_for_search(keyword)
        if not keyword_normal:
            return

        for page_index, page_text in enumerate(self.pages):
            display_text = self.add_cjk_break_opportunities(page_text)
            normalized_text, _ = self.make_normalized_index_map(display_text)

            start = 0
            occurrence = 0
            while True:
                pos = normalized_text.find(keyword_normal, start)
                if pos < 0:
                    break
                self.search_results.append({
                    "page": page_index,
                    "keyword": keyword,
                    "occurrence": occurrence
                })
                occurrence += 1
                start = pos + max(len(keyword_normal), 1)

    def search_next(self):
        keyword = self.search_edit.text().strip() if hasattr(self, "search_edit") else ""
        if not keyword or self.is_pdf:
            return

        if keyword != self.search_keyword or not self.search_results:
            self.search_keyword = keyword
            self.build_search_results(keyword)

        if not self.search_results:
            self.status.showMessage("검색 결과가 없습니다.", 3000)
            return

        self.current_search_index = (self.current_search_index + 1) % len(self.search_results)
        self.go_search_result()

    def search_prev(self):
        keyword = self.search_edit.text().strip() if hasattr(self, "search_edit") else ""
        if not keyword or self.is_pdf:
            return

        if keyword != self.search_keyword or not self.search_results:
            self.search_keyword = keyword
            self.build_search_results(keyword)

        if not self.search_results:
            self.status.showMessage("검색 결과가 없습니다.", 3000)
            return

        self.current_search_index = (self.current_search_index - 1) % len(self.search_results)
        self.go_search_result()

    def go_search_result(self):
        if not self.search_results:
            return

        result = self.search_results[self.current_search_index]
        page_index = result["page"]

        if self.view_mode == 2:
            self.current_page = page_index
            if self.current_page % 2 == 1:
                self.current_page -= 1
        else:
            self.current_page = page_index

        self.display_page()
        QTimer.singleShot(120, lambda r=result: self.select_search_text(r))

        self.status.showMessage(f"검색 결과 {self.current_search_index + 1} / {len(self.search_results)}", 3000)

    def select_search_text(self, result):
        page_index = result.get("page")
        keyword = result.get("keyword", "")
        occurrence = result.get("occurrence", 0)

        if page_index is None or not keyword:
            return

        if self.view_mode == 1:
            viewer = self.right_viewer
        else:
            if page_index == self.current_page:
                viewer = self.left_viewer
            elif page_index == self.current_page + 1:
                viewer = self.right_viewer
            else:
                return

        plain_text = viewer.toPlainText()
        normalized_text, index_map = self.make_normalized_index_map(plain_text)
        keyword_normal = self.normalize_for_search(keyword)

        if not keyword_normal:
            return

        start = 0
        found_pos = -1
        for _ in range(occurrence + 1):
            found_pos = normalized_text.find(keyword_normal, start)
            if found_pos < 0:
                return
            start = found_pos + max(len(keyword_normal), 1)

        if found_pos < 0 or found_pos >= len(index_map):
            return

        start_pos = index_map[found_pos]
        end_normal_pos = found_pos + len(keyword_normal) - 1
        if end_normal_pos >= len(index_map):
            end_pos = len(plain_text)
        else:
            end_pos = index_map[end_normal_pos] + 1

        start_pos = max(0, min(start_pos, len(plain_text)))
        end_pos = max(start_pos, min(end_pos, len(plain_text)))

        cursor = viewer.textCursor()
        cursor.setPosition(start_pos)
        cursor.setPosition(end_pos, QTextCursor.MoveMode.KeepAnchor)
        viewer.setTextCursor(cursor)
        viewer.setFocus()
        viewer.ensureCursorVisible()

    # ---------------- Annotation ----------------
    def get_active_viewer_and_page(self):
        if self.is_pdf:
            return None, self.current_page

        if self.view_mode == 1:
            return self.right_viewer, self.current_page

        if self.right_viewer.hasFocus():
            return self.right_viewer, self.current_page + 1
        return self.left_viewer, self.current_page

    def apply_annotations_to_viewer(self, viewer, page_index):
        for ann in self.annotations:
            if ann.get("page") != page_index:
                continue

            text = viewer.toPlainText()
            start = min(ann.get("start", 0), len(text))
            end = min(ann.get("end", 0), len(text))

            cursor = viewer.textCursor()
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)

            fmt = QTextCharFormat()
            if ann.get("highlight"):
                fmt.setBackground(QColor(ann.get("color", "#ffff00")))
            if ann.get("underline"):
                fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SingleUnderline)
                fmt.setUnderlineColor(QColor(ann.get("underline_color", "#1d4ed8")))
            cursor.mergeCharFormat(fmt)

    def save_annotation(self, highlight=False, underline=False):
        viewer, page_index = self.get_active_viewer_and_page()
        if viewer is None:
            return
        cursor = viewer.textCursor()
        if not cursor.hasSelection():
            return

        ann = {
            "page": page_index,
            "start": cursor.selectionStart(),
            "end": cursor.selectionEnd(),
            "highlight": highlight,
            "underline": underline,
            "color": self.highlight_color.name(),
            "underline_color": self.underline_color.name(),
            "text": cursor.selectedText()
        }
        self.annotations.append(ann)
        self.save_state()

    def highlight_text(self):
        viewer, _ = self.get_active_viewer_and_page()
        if viewer is None:
            return
        cursor = viewer.textCursor()
        if not cursor.hasSelection():
            return
        fmt = QTextCharFormat()
        fmt.setBackground(self.highlight_color)
        cursor.mergeCharFormat(fmt)
        self.save_annotation(highlight=True)

    def underline_text(self):
        viewer, _ = self.get_active_viewer_and_page()
        if viewer is None:
            return
        cursor = viewer.textCursor()
        if not cursor.hasSelection():
            return
        fmt = QTextCharFormat()
        fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SingleUnderline)
        fmt.setUnderlineColor(self.underline_color)
        cursor.mergeCharFormat(fmt)
        self.save_annotation(underline=True)

    def choose_highlight_color(self):
        color = QColorDialog.getColor(self.highlight_color, self)
        if color.isValid():
            self.highlight_color = color
            self.save_app_config()

    def choose_underline_color(self):
        color = QColorDialog.getColor(self.underline_color, self)
        if color.isValid():
            self.underline_color = color
            self.save_app_config()

    def clear_current_page_annotations(self):
        pages_to_clear = [self.current_page]
        if self.view_mode == 2:
            pages_to_clear.append(self.current_page + 1)
        self.annotations = [ann for ann in self.annotations if ann.get("page") not in pages_to_clear]
        self.save_state()
        self.display_page()

    # ---------------- Bookmark ----------------
    def set_bookmark(self):
        total = self.total_pages()
        if total <= 0:
            return
        self.bookmark_page = self.current_page
        self.bookmark_side = "right" if self.view_mode == 1 else "left"
        self.save_state()
        self.display_page()

    def set_bookmark_left(self):
        total = self.total_pages()
        if total <= 0:
            return
        if self.view_mode == 2:
            self.bookmark_page = self.current_page
            self.bookmark_side = "left"
        else:
            self.bookmark_page = self.current_page
            self.bookmark_side = "right"
        self.save_state()
        self.display_page()

    def set_bookmark_right(self):
        total = self.total_pages()
        if total <= 0:
            return
        if self.view_mode == 2 and self.current_page + 1 < total:
            self.bookmark_page = self.current_page + 1
            self.bookmark_side = "right"
        else:
            self.bookmark_page = self.current_page
            self.bookmark_side = "right"
        self.save_state()
        self.display_page()

    def go_bookmark(self):
        if self.bookmark_page is None:
            return
        total = self.total_pages()
        if total <= 0:
            return

        self.current_page = min(self.bookmark_page, total - 1)
        if self.view_mode == 2 and self.current_page % 2 == 1:
            self.current_page -= 1
        self.display_page()
        self.save_state()

    def clear_bookmark(self):
        self.bookmark_page = None
        self.bookmark_side = "right"
        self.save_state()
        self.display_page()

    # ---------------- Navigation / options ----------------
    def total_pages(self):
        return self.pdf_page_count if self.is_pdf else len(self.pages)

    def next_page(self):
        total = self.total_pages()
        if total <= 0:
            return
        step = 1 if self.view_mode == 1 else 2
        if self.current_page + step < total:
            self.current_page += step
            self.display_page()
            self.save_state()

    def prev_page(self):
        total = self.total_pages()
        if total <= 0:
            return
        step = 1 if self.view_mode == 1 else 2
        if self.current_page - step >= 0:
            self.current_page -= step
            self.display_page()
            self.save_state()

    def handle_edge_click(self, direction):
        self.next_page() if direction == "next" else self.prev_page()

    def keyPressEvent(self, a0):
        if a0 is None:
            return

        key = a0.key()

        if key in [
            Qt.Key.Key_Right, Qt.Key.Key_Down,
            Qt.Key.Key_PageDown, Qt.Key.Key_Space
        ]:
            self.next_page()
            return

        if key in [
            Qt.Key.Key_Left, Qt.Key.Key_Up,
            Qt.Key.Key_PageUp, Qt.Key.Key_Backspace
        ]:
            self.prev_page()
            return

        if key == Qt.Key.Key_F11:
            self.toggle_fullscreen()
            return

        if key == Qt.Key.Key_Escape and self.isFullScreen():
            self.showNormal()
            return

        super().keyPressEvent(a0)

    def resizeEvent(self, a0):
        if a0 is None:
            return
        super().resizeEvent(a0)

        if hasattr(self, "_resize_timer"):
            self._resize_timer.stop()
        else:
            self._resize_timer = QTimer(self)
            self._resize_timer.setSingleShot(True)
            self._resize_timer.timeout.connect(self._on_resize_finished)

        self._resize_timer.start(250)

    def _on_resize_finished(self):
        self.update_page_geometry()
        if self.is_pdf:
            self.display_page()
        elif self.pages:
            self.repaginate()
            self.display_page()
            self.update_page_thumbnails()
        else:
            self.display_page()

    def set_view_mode_from_menu(self, mode):
        self.view_mode = mode
        if self.view_mode == 2 and self.current_page % 2 == 1:
            self.current_page -= 1
        self.update_page_geometry()
        if not self.is_pdf and self.pages:
            self.repaginate()
        self.display_page()
        self.update_page_thumbnails()
        if self.view_mode == 2:
            QTimer.singleShot(0, self.force_initial_two_page_layout)
        self.save_app_config()
        self.save_state()

    def change_view_mode(self, index=None):
        if not hasattr(self, "view_combo"):
            return
        self.view_mode = 1 if self.view_combo.currentIndex() == 0 else 2
        self.set_view_mode_from_menu(self.view_mode)

    def toggle_fit_to_window(self, checked):
        self.fit_to_window = checked
        self.update_page_geometry()
        self.display_page()
        self.update_page_thumbnails()
        self.save_app_config()

    def toggle_page_thumbnail_panel(self, checked):
        self.page_thumbnail_visible = bool(checked)
        if hasattr(self, "page_thumb_panel"):
            self.page_thumb_panel.setVisible(self.page_thumbnail_visible)
        if hasattr(self, "thumb_panel_action"):
            self.thumb_panel_action.setChecked(self.page_thumbnail_visible)
        if hasattr(self, "thumb_ribbon_action"):
            self.thumb_ribbon_action.setChecked(self.page_thumbnail_visible)
        QTimer.singleShot(50, self._on_resize_finished)
        self.save_app_config()

    def toggle_fullscreen(self):
        self.showNormal() if self.isFullScreen() else self.showFullScreen()

    def toggle_font_bold(self, checked):
        self.font_bold = checked
        if hasattr(self, "bold_action"):
            self.bold_action.setChecked(checked)
        if hasattr(self, "bold_menu_action"):
            self.bold_menu_action.setChecked(checked)
        self.apply_reader_font()
        if not self.is_pdf and self.full_text:
            self.repaginate()
        self.display_page()
        self.update_page_thumbnails()
        self.save_app_config()
        self.save_state()

    def change_paper_size(self, name):
        self.paper_size = name
        self.update_page_geometry()
        if not self.is_pdf and self.pages:
            self.repaginate()
        self.display_page()
        self.update_page_thumbnails()
        self.save_app_config()

    def change_inner_margin_x(self, value):
        self.inner_margin_x = value
        self.refresh_text_layout()

    def change_inner_margin_y(self, value):
        self.inner_margin_y = value
        self.refresh_text_layout()

    def change_paragraph_spacing(self, value):
        self.paragraph_spacing = value
        self.refresh_text_layout()

    def change_first_line_indent(self, value):
        self.first_line_indent = value
        self.refresh_text_layout()

    def change_line_height(self, value):
        self.line_height = float(value)
        self.refresh_text_layout()

    def refresh_text_layout(self):
        self.apply_text_style()
        if not self.is_pdf and self.pages:
            self.repaginate()
        self.display_page()
        self.update_page_thumbnails()
        self.save_app_config()

    def change_font(self, name):
        if self.loading_state:
            return
        self.app_font_name = name
        self.apply_reader_font()
        if not self.is_pdf and self.full_text:
            self.repaginate()
        self.display_page()
        self.update_page_thumbnails()
        self.save_app_config()

    def change_font_size(self, size):
        if self.loading_state:
            return
        self.app_font_size = size
        self.apply_reader_font()
        if not self.is_pdf and self.full_text:
            self.repaginate()
        self.display_page()
        self.update_page_thumbnails()
        self.save_app_config()

    def change_paper_color_without_save(self, name):
        self.paper_theme = name
        if name == "화이트":
            paper, text, outer = "#ffffff", "#000000", "#d7d7d7"
        elif name == "블랙":
            paper, text, outer = "#111111", "#ffffff", "#000000"
        else:
            paper, text, outer = "#fbf6e8", "#2b2118", "#3a3028"

        self.canvas.set_theme(paper, text, outer)
        self.apply_text_style()

    def change_paper_color(self, name):
        self.change_paper_color_without_save(name)
        self.update_page_thumbnails()
        self.save_app_config()

    def set_ui_font_size_from_menu(self, size):
        self.ui_font_size = size
        if hasattr(self, "ui_font_action_group"):
            for action in self.ui_font_action_group.actions():
                action.setChecked(action.text() == f"{size} pt")
        self.apply_ui_font_size()
        self.save_app_config()

    def apply_global_ui_config_to_widgets(self):
        self.loading_state = True
        try:
            self.apply_reader_font()
            if hasattr(self, "page_thumb_panel"):
                self.page_thumb_panel.setVisible(self.page_thumbnail_visible)
            self.change_paper_color_without_save(self.paper_theme)
        finally:
            self.loading_state = False

    # ---------------- Recent / persistence ----------------
    def add_recent_book(self, file_path):
        if not file_path:
            return
        file_path = os.path.abspath(file_path)
        self.recent_books = [p for p in self.recent_books if os.path.abspath(p) != file_path]
        self.recent_books.insert(0, file_path)
        self.recent_books = self.recent_books[:self.max_recent_books]
        self.update_recent_books_menu()
        self.save_app_config()

    def update_recent_books_menu(self):
        if not hasattr(self, "recent_menu"):
            return
        self.recent_menu.clear()
        self.recent_books = [p for p in self.recent_books if os.path.exists(p)][:self.max_recent_books]
        if not self.recent_books:
            a = QAction("최근 항목 없음", self)
            a.setEnabled(False)
            self.recent_menu.addAction(a)
            return
        for path in self.recent_books:
            self.recent_menu.addAction(self.action(os.path.basename(path), lambda checked=False, p=path: self.open_recent_book(p)))
        self.recent_menu.addSeparator()
        self.recent_menu.addAction(self.action("최근 목록 지우기", self.clear_recent_books))

    def open_recent_book(self, file_path):
        if file_path and os.path.exists(file_path):
            self.open_document_file(file_path)
        else:
            self.recent_books = [p for p in self.recent_books if p != file_path]
            self.update_recent_books_menu()
            self.save_app_config()
            self.status.showMessage("파일을 찾을 수 없습니다.", 3000)

    def clear_recent_books(self):
        self.recent_books = []
        self.update_recent_books_menu()
        self.save_app_config()

    def state_file_path(self):
        return os.path.join(self.data_dir, f"{self.book_id}.json")

    def save_state(self):
        if self.loading_state or not self.current_file or not self.book_id:
            return

        elapsed = int(time.time() - self.start_time)
        total_read_seconds = self.saved_read_seconds + elapsed

        data = {
            "file": self.current_file,
            "current_page": self.current_page,
            "bookmark_page": self.bookmark_page,
            "bookmark_side": self.bookmark_side,
            "read_seconds": total_read_seconds,
            "annotations": self.annotations,
            "last_opened": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        try:
            with open(self.state_file_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.saved_read_seconds = total_read_seconds
            self.start_time = time.time()
        except Exception as e:
            print("상태 저장 오류:", e)

    def load_state(self):
        path = self.state_file_path()
        if not os.path.exists(path):
            return
        self.loading_state = True
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            total = self.total_pages()
            self.bookmark_page = data.get("bookmark_page")
            self.bookmark_side = data.get("bookmark_side", "right")
            self.saved_read_seconds = data.get("read_seconds", 0)
            self.annotations = data.get("annotations", [])
            self.current_page = min(data.get("current_page", 0), max(total - 1, 0))
            if self.view_mode == 2 and self.current_page % 2 == 1:
                self.current_page -= 1
            self.start_time = time.time()
        except Exception as e:
            print("상태 불러오기 오류:", e)
        finally:
            self.loading_state = False

    def load_app_config(self):
        if not os.path.exists(self.app_config_file):
            return
        try:
            with open(self.app_config_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.last_open_dir = data.get("last_open_dir", self.last_open_dir)
            self.recent_books = data.get("recent_books", self.recent_books)
            self.view_mode = data.get("view_mode", self.view_mode)
            self.paper_size = data.get("paper_size", self.paper_size)
            self.fit_to_window = data.get("fit_to_window", self.fit_to_window)
            self.inner_margin_x = data.get("inner_margin_x", self.inner_margin_x)
            self.inner_margin_y = data.get("inner_margin_y", self.inner_margin_y)
            self.paragraph_spacing = data.get("paragraph_spacing", self.paragraph_spacing)
            self.first_line_indent = data.get("first_line_indent", self.first_line_indent)
            self.line_height = data.get("line_height", self.line_height)
            self.font_bold = data.get("font_bold", self.font_bold)
            self.ui_font_size = data.get("ui_font_size", self.ui_font_size)
            self.page_thumbnail_visible = data.get("page_thumbnail_visible", self.page_thumbnail_visible)
            self.paper_theme = data.get("paper_theme", self.paper_theme)
            self.app_font_name = data.get("font", self.app_font_name)
            self.app_font_size = data.get("font_size", self.app_font_size)
            self.highlight_color = QColor(data.get("highlight_color", "#ffff00"))
            self.underline_color = QColor(data.get("underline_color", "#1d4ed8"))
        except Exception as e:
            print("앱 설정 불러오기 오류:", e)

    def save_app_config(self):
        data = {
            "last_open_dir": self.last_open_dir,
            "recent_books": self.recent_books,
            "view_mode": self.view_mode,
            "paper_size": self.paper_size,
            "fit_to_window": self.fit_to_window,
            "inner_margin_x": self.inner_margin_x,
            "inner_margin_y": self.inner_margin_y,
            "paragraph_spacing": self.paragraph_spacing,
            "first_line_indent": self.first_line_indent,
            "line_height": self.line_height,
            "font": self.app_font_name,
            "font_size": self.app_font_size,
            "font_bold": self.font_bold,
            "ui_font_size": self.ui_font_size,
            "page_thumbnail_visible": self.page_thumbnail_visible,
            "paper_theme": self.paper_theme,
            "highlight_color": self.highlight_color.name(),
            "underline_color": self.underline_color.name(),
        }
        try:
            with open(self.app_config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("앱 설정 저장 오류:", e)

    # ---------------- Misc ----------------
    def get_system_fonts(self):
        families = QFontDatabase.families()
        preferred = [
            "맑은 고딕", "Malgun Gothic", "나눔고딕", "NanumGothic",
            "Noto Sans CJK KR", "Noto Serif CJK KR",
            "Apple SD Gothic Neo", "Arial", "Times New Roman", "Georgia"
        ]
        result = []
        for font in preferred:
            if font in families and font not in result:
                result.append(font)
        for font in families:
            if font not in result:
                result.append(font)
        return result

    def update_view_mode(self):
        self.canvas.set_view_mode(self.view_mode)

        if self.bookmark_page is None:
            self.canvas.set_bookmark(False)
            return

        if self.view_mode == 1:
            self.canvas.set_bookmark(self.bookmark_page == self.current_page, "right")
        else:
            if self.bookmark_page == self.current_page:
                self.canvas.set_bookmark(True, "left")
            elif self.bookmark_page == self.current_page + 1:
                self.canvas.set_bookmark(True, "right")
            else:
                self.canvas.set_bookmark(False)

    def update_canvas_page_numbers(self):
        total = self.total_pages()
        if total <= 0:
            self.canvas.set_page_numbers([])
            return

        numbers = []
        if self.view_mode == 1:
            numbers.append({"rect_index": 0, "text": str(self.current_page + 1), "align": "center"})
        else:
            left_page = self.current_page
            right_page = self.current_page + 1
            numbers.append({"rect_index": 0, "text": str(left_page + 1), "align": "left"})
            if right_page < total:
                numbers.append({"rect_index": 1, "text": str(right_page + 1), "align": "right"})
        self.canvas.set_page_numbers(numbers if not self.is_pdf else [])

    def update_log(self):
        total = self.total_pages()
        current = self.current_page + 1 if total else 0
        progress = int((current / total) * 100) if total else 0
        bookmark_text = ""
        if self.bookmark_page is not None:
            bookmark_text = f" / 책갈피: {self.bookmark_page + 1}"
        self.page_label.setText(f"페이지: {current} / {total}{bookmark_text}")
        self.progress_label.setText(f"진행율: {progress}%")

    def update_time_log(self):
        self.time_label.setText("현재 시간: " + datetime.now().strftime("%H:%M:%S"))
        elapsed = int(time.time() - self.start_time)
        total = self.saved_read_seconds + elapsed
        h, m, s = total // 3600, total % 3600 // 60, total % 60
        self.read_time_label.setText(f"읽은 시간: {h:02d}:{m:02d}:{s:02d}")

    def show_edit_context_menu(self, pos):
        if hasattr(self, "context_menu"):
            self.context_menu.popup(pos)

    def closeEvent(self, a0):
        if a0 is None:
            return
        self.save_state()
        self.save_app_config()
        self.close_pdf()

        if a0 is not None:
            a0.accept()

def main():
    app = QApplication(sys.argv)
    reader = DocumentReader()
    reader.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
