import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QFileDialog, QGraphicsView, 
                             QGraphicsScene, QGraphicsPixmapItem, QMenuBar, 
                             QGraphicsLineItem, QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem,
                             QGraphicsTextItem, QPushButton, QComboBox, QSlider, QCheckBox, QScrollArea, QMessageBox, QGridLayout)
from PySide6.QtGui import QAction, QPixmap, QPen, QColor, QPainter, QImage, QIcon, QFont, QPainterPath
from PySide6.QtCore import Qt, QRectF, QSize, QPointF

# 1. 마우스 조작 가능 이미지 아이템
class DraggablePixmapItem(QGraphicsPixmapItem):
    def __init__(self, pixmap):
        super().__init__(pixmap)
        self.setFlag(QGraphicsPixmapItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsPixmapItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)

# 2. 이미지 편집 뷰 영역
class PhotoEditView(QGraphicsView):
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setBackgroundBrush(QColor(26, 26, 26))
        
        self.image_item = None
        self.current_rotation = 0
        
        self.undo_stack = []
        self.redo_stack = []

    def set_image(self, pixmap):
        if self.image_item:
            self.scene().removeItem(self.image_item)
        
        self.image_item = DraggablePixmapItem(pixmap)
        self.image_item.setZValue(0)
        self.scene().addItem(self.image_item)
        
        self.image_item.setTransformOriginPoint(self.image_item.pixmap().width() / 2, self.image_item.pixmap().height() / 2)
        self.image_item.setPos(0, 0)
        self.current_rotation = 0
        self.image_item.setRotation(0)
        self.image_item.setScale(1.0)
        
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.save_current_state()

    def zoom_in(self):
        if self.image_item:
            self.save_current_state()
            self.image_item.setScale(self.image_item.scale() * 1.1)

    def zoom_out(self):
        if self.image_item:
            self.save_current_state()
            self.image_item.setScale(self.image_item.scale() * 0.9)

    def wheelEvent(self, event):
        if not self.image_item:
            return
        self.save_current_state()
        if event.angleDelta().y() > 0:
            self.image_item.setScale(self.image_item.scale() * 1.1)
        else:
            self.image_item.setScale(self.image_item.scale() * 0.9)

    def rotate_image(self, angle):
        if not self.image_item:
            return
        self.save_current_state()
        self.current_rotation = (self.current_rotation + angle) % 360
        self.image_item.setRotation(self.current_rotation)

    def clear_image(self):
        if self.image_item:
            self.save_current_state()
            self.scene().removeItem(self.image_item)
            self.image_item = None
            self.undo_stack.clear()
            self.redo_stack.clear()

    def save_current_state(self):
        if not self.image_item:
            return
        state = {
            'pos': QPointF(self.image_item.pos()),
            'scale': self.image_item.scale(),
            'rotation': self.image_item.rotation()
        }
        self.undo_stack.append(state)
        self.redo_stack.clear()

    def undo(self):
        if len(self.undo_stack) > 1:
            current_state = {
                'pos': QPointF(self.image_item.pos()),
                'scale': self.image_item.scale(),
                'rotation': self.image_item.rotation()
            }
            self.redo_stack.append(current_state)
            prev_state = self.undo_stack.pop()
            
            if prev_state['pos'] == current_state['pos'] and prev_state['scale'] == current_state['scale'] and prev_state['rotation'] == current_state['rotation']:
                if self.undo_stack:
                    prev_state = self.undo_stack.pop()
                else:
                    return

            self.image_item.setPos(prev_state['pos'])
            self.image_item.setScale(prev_state['scale'])
            self.image_item.setRotation(prev_state['rotation'])
            self.current_rotation = prev_state['rotation']

    def redo(self):
        if self.redo_stack:
            next_state = self.redo_stack.pop()
            current_state = {
                'pos': QPointF(self.image_item.pos()),
                'scale': self.image_item.scale(),
                'rotation': self.image_item.rotation()
            }
            self.undo_stack.append(current_state)
            
            self.image_item.setPos(next_state['pos'])
            self.image_item.setScale(next_state['scale'])
            self.image_item.setRotation(next_state['rotation'])
            self.current_rotation = next_state['rotation']


# 3. 사진 유형용 커스텀 그리드 아이템 위젯
class SampleCardWidget(QWidget):
    def __init__(self, name, full_key, parent_app):
        super().__init__()
        self.full_key = full_key
        self.parent_app = parent_app
        self.is_selected = False
        self.init_ui(name)
        
    def init_ui(self, name):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.img_label = QLabel()
        self.img_label.setFixedSize(110, 130)
        self.img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        sample_thumb = QPixmap(110, 130)
        sample_thumb.fill(QColor(45, 45, 45))
        painter = QPainter(sample_thumb)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        if "증멍증냥" in name:
            painter.setBrush(QColor(140, 110, 90))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(25, 40, 60, 60)
            painter.drawEllipse(20, 25, 25, 25)
            painter.drawEllipse(65, 25, 25, 25)
        else:
            painter.setBrush(QColor(80, 85, 95))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(10, 85, 90, 80)
            painter.setBrush(QColor(240, 205, 180))
            painter.drawEllipse(27, 25, 56, 65)
            painter.setBrush(QColor(55, 50, 50))
            painter.drawEllipse(25, 15, 60, 35)
            
        painter.end()
        self.img_label.setPixmap(sample_thumb)
        
        self.txt_label = QLabel(name)
        self.txt_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.txt_label.setWordWrap(True)
        self.txt_label.setFixedWidth(110)

        self.font_size = 12
        self.txt_label.setStyleSheet(f"color: #a5a7aa; font-size: {self.font_size}px; font-weight: normal; background: transparent;")
        
        layout.addWidget(self.img_label)
        layout.addWidget(self.txt_label)
        
        self.setObjectName("SampleCard")
        self.update_style()

    def update_style(self):
        if self.is_selected:
            self.setStyleSheet("QWidget#SampleCard { background-color: #222428; border: 2px solid #1a73e8; border-radius: 6px; }")
            self.txt_label.setStyleSheet(f"color: #ffffff; font-size: {self.font_size}px; font-weight: bold; background: transparent;")
        else:
            self.setStyleSheet("QWidget#SampleCard { background-color: #151619; border: 1px solid #25272c; border-radius: 6px; } QWidget#SampleCard:hover { background-color: #1c1e22; border: 1px solid #3a3f47; }")
            self.txt_label.setStyleSheet(f"color: #a5a7aa; font-size: {self.font_size}px; font-weight: normal; background: transparent;")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.parent_app.select_preset_card(self.full_key)


# 4. 메인 윈도우 클래스
class PhotoProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ID Photo Master PRO v1.5")
        
        self.size_presets = {
            "여권 (3.5cm x 4.5cm)": (413, 531),
            "운전면허 (3.5cm x 4.5cm)": (413, 531),
            "주민등록 (3.5cm x 4.5cm)": (413, 531),
            "반명함 (3.0cm x 4.0cm)": (354, 472),
            "명함 (5.0cm x 7.0cm)": (591, 827),
            "미.중.일 비자 (공통규격)": (591, 591),
            "인화용 (10.2cm x 15.2cm)": (1205, 1795),
            "증멍증냥 (반려동물 규격)": (413, 413)
        }
        self.current_preset_name = "여권 (3.5cm x 4.5cm)"
        self.current_target_size = self.size_presets[self.current_preset_name]
        self.card_widgets = {}
        
        self.init_ui()
        self.apply_image_style_theme()
        self.showMaximized()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)
        
        # [LEFT PANEL]
        left_container = QWidget()
        left_container.setMinimumWidth(280)
        left_container.setMaximumWidth(320)
        left_vbox = QVBoxLayout(left_container)
        left_vbox.setContentsMargins(0, 0, 0, 0)
        
        left_title = QLabel("사진 유형 샘플")
        left_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #ffffff; padding-left: 5px; margin-bottom: 5px;")
        left_vbox.addWidget(left_title)
        
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setStyleSheet("background-color: #151619; border: 1px solid #25272c; border-radius: 4px;")
        
        grid_widget = QWidget()
        grid_widget.setStyleSheet("background-color: #151619;")
        self.grid_layout = QGridLayout(grid_widget)
        self.grid_layout.setContentsMargins(8, 8, 8, 8)
        self.grid_layout.setSpacing(8)
        
        row, col = 0, 0
        for name in self.size_presets.keys():
            clean_name = name.split(" ")[0]
            dimensions = f"{self.size_presets[name][0]}x{self.size_presets[name][1]} px"
            display_text = f"{clean_name}\n({dimensions})"
            
            card = SampleCardWidget(display_text, name, self)
            self.card_widgets[name] = card
            
            self.grid_layout.addWidget(card, row, col, Qt.AlignmentFlag.AlignCenter)
            col += 1
            if col > 1:
                col = 0
                row += 1
                
        self.grid_layout.setRowStretch(row + 1, 1)
        left_scroll.setWidget(grid_widget)
        left_vbox.addWidget(left_scroll)
        
        # [CENTER PANEL]
        center_panel = QWidget()
        center_layout = QVBoxLayout(center_panel)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(5)
        
        self.scene = QGraphicsScene()
        self.view = PhotoEditView(self.scene)
        
        # --- 상단 미니 툴바 영역 (90도 회전 아이콘 추가) ---
        mini_toolbar = QHBoxLayout()
        mini_toolbar.setContentsMargins(5, 0, 5, 0)
        
        btn_close = QPushButton("✕")
        btn_close.clicked.connect(self.view.clear_image)
        btn_zoom_in = QPushButton("🔍+")
        btn_zoom_in.clicked.connect(self.view.zoom_in)
        btn_zoom_out = QPushButton("🔍-")
        btn_zoom_out.clicked.connect(self.view.zoom_out)
        
        # [추가] 반시계 방향 90도 / 시계 방향 90도 회전 기능 아이콘
        btn_rot_ccw90 = QPushButton("↺")
        btn_rot_ccw90.setToolTip("반시계 방향 90도 회전")
        btn_rot_ccw90.clicked.connect(lambda: self.view.rotate_image(-90))
        
        btn_rot_cw90 = QPushButton("↻")
        btn_rot_cw90.setToolTip("시계 방향 90도 회전")
        btn_rot_cw90.clicked.connect(lambda: self.view.rotate_image(90))
        
        btn_undo = QPushButton("↶")
        btn_undo.clicked.connect(self.view.undo)
        btn_redo = QPushButton("↷")
        btn_redo.clicked.connect(self.view.redo)
        
        # 툴바 아이템 순서대로 배치
        toolbar_buttons = [btn_close, btn_zoom_in, btn_zoom_out, btn_rot_ccw90, btn_rot_cw90, btn_undo, btn_redo]
        for btn in toolbar_buttons:
            btn.setFixedSize(30, 30)
            btn.setStyleSheet("background: transparent; color: #888888; border: none; font-size: 15px; font-weight: bold;")
            mini_toolbar.addWidget(btn)
        
        mini_toolbar.addStretch()
        
        btn_help = QPushButton("❓")
        btn_help.clicked.connect(self.show_help_dialog)
        btn_setting = QPushButton("⚙")
        btn_setting.clicked.connect(self.show_settings_dialog)
        
        for btn in [btn_help, btn_setting]:
            btn.setFixedSize(30, 30)
            btn.setStyleSheet("background: transparent; color: #888888; border: none; font-size: 14px;")
            mini_toolbar.addWidget(btn)
            
        center_layout.addLayout(mini_toolbar)
        center_layout.addWidget(self.view)
        
        # [RIGHT PANEL]
        scroll_area = QScrollArea()
        scroll_area.setObjectName("RightScrollArea")
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumWidth(260)
        scroll_area.setMaximumWidth(340)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        right_panel = QWidget()
        right_panel.setObjectName("RightPanel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(15, 15, 15, 15)
        right_layout.setSpacing(12)
        
        right_layout.addWidget(QLabel("사진 규격 선택"))
        self.combo_presets = QComboBox()
        for k in self.size_presets.keys():
            self.combo_presets.addItem(k)
        self.combo_presets.currentTextChanged.connect(self.change_photo_size)
        right_layout.addWidget(self.combo_presets)
        
        shortcut_sizes = ["여권", "반명함", "비자", "증멍증냥"]
        for s_name in shortcut_sizes:
            s_btn = QPushButton(s_name)
            s_btn.setObjectName("ShortcutBtn")
            full_key = [k for k in self.size_presets.keys() if s_name in k]
            if full_key:
                s_btn.clicked.connect(lambda chk=False, nk=full_key[0]: self.change_photo_size(nk))
            right_layout.addWidget(s_btn)
            
        right_layout.addWidget(QLabel("얼굴 중심 및 비율"))
        self.chk_guide = QCheckBox("가이드라인 표시")
        self.chk_guide.setChecked(True)
        self.chk_guide.stateChanged.connect(self.toggle_guides)
        right_layout.addWidget(self.chk_guide)
        
        slider_label_layout = QHBoxLayout()
        slider_label_layout.addWidget(QLabel("정수리/턱 비율 표준"))
        self.lbl_ratio_val = QLabel("0")
        self.lbl_ratio_val.setAlignment(Qt.AlignmentFlag.AlignRight)
        slider_label_layout.addWidget(self.lbl_ratio_val)
        right_layout.addLayout(slider_label_layout)
        
        self.ratio_slider = QSlider(Qt.Orientation.Horizontal)
        self.ratio_slider.setRange(-50, 50)
        self.ratio_slider.setValue(0)
        self.ratio_slider.valueChanged.connect(lambda val: self.lbl_ratio_val.setText(str(val)))
        right_layout.addWidget(self.ratio_slider)
        
        self.btn_rotate_cw = QPushButton("미세 회전 (시계방향 5°)")
        self.btn_rotate_cw.setObjectName("ActionBlueBtn")
        self.btn_rotate_cw.clicked.connect(lambda: self.view.rotate_image(5))
        right_layout.addWidget(self.btn_rotate_cw)
        
        right_layout.addWidget(QLabel("파일 형식 저장"))
        right_layout.addWidget(QLabel("저장 파일 형식"))
        self.combo_formats = QComboBox()
        self.combo_formats.addItems(["JPEG (*.jpg)", "JPEG Alternative (*.jpeg)", "PNG (*.png)", "TIFF (*.tiff)"])
        right_layout.addWidget(self.combo_formats)
        
        self.btn_save = QPushButton("저장하기")
        self.btn_save.setObjectName("ActionBlueBtn")
        self.btn_save.clicked.connect(self.save_image)
        right_layout.addWidget(self.btn_save)
        
        right_layout.addStretch()
        
        self.lbl_status = QLabel("300 DPI - ready")
        self.lbl_status.setObjectName("StatusLabel")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignRight)
        right_layout.addWidget(self.lbl_status)
        
        scroll_area.setWidget(right_panel)
        
        main_layout.addWidget(left_container)
        main_layout.addWidget(center_panel, stretch=70)
        main_layout.addWidget(scroll_area, stretch=30)
        
        self.create_menu_bar()
        self.select_preset_card(self.current_preset_name)
        self.update_guide_lines()

    def create_menu_bar(self):
        menu_bar = self.menuBar()
        menu_bar.clear()
        file_menu = menu_bar.addMenu("File")
        file_menu.addAction(QAction("Open Image...", self, shortcut="Ctrl+O", triggered=self.open_image))
        file_menu.addAction(QAction("Save Image...", self, shortcut="Ctrl+S", triggered=self.save_image))
        file_menu.addSeparator()
        file_menu.addAction(QAction("Exit App", self, shortcut="Ctrl+Q", triggered=self.close))
        
        edit_menu = menu_bar.addMenu("Edit")
        edit_menu.addAction(QAction("Undo Move/Zoom", self, shortcut="Ctrl+Z", triggered=self.view.undo))
        edit_menu.addAction(QAction("Redo Move/Zoom", self, shortcut="Ctrl+Y", triggered=self.view.redo))

    def show_help_dialog(self):
        QMessageBox.information(self, "Help & About", "<h3>ID Photo Master PRO v1.5</h3><p>국가 규격 가이드라인을 완벽 지원하는 사진 편집 솔루션입니다.</p>")

    def show_settings_dialog(self):
        QMessageBox.information(self, "Preferences", "해상도는 국제 표준 300 DPI 인쇄 설계규격 기준입니다.")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.fit_guide_in_view()

    def select_preset_card(self, size_name):
        if size_name not in self.size_presets:
            return
        self.current_preset_name = size_name
        self.current_target_size = self.size_presets[size_name]
        
        for k, card in self.card_widgets.items():
            card.is_selected = (k == size_name)
            card.update_style()
            
        if self.combo_presets.currentText() != size_name:
            self.combo_presets.setCurrentText(size_name)
            
        self.update_guide_lines()

    def change_photo_size(self, size_name):
        self.select_preset_card(size_name)

    def toggle_guides(self, state):
        visible = (state == 2 or state == True)
        for item in self.scene.items():
            if isinstance(item, (QGraphicsLineItem, QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem, QGraphicsTextItem)):
                item.setVisible(visible)

    def update_guide_lines(self):
        for item in self.scene.items():
            if isinstance(item, (QGraphicsLineItem, QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem, QGraphicsTextItem)):
                self.scene.removeItem(item)

        w, h = self.current_target_size
        self.scene.setSceneRect(-130, -50, w + 260, h + 100)
        
        # 1. 외곽 초록색 크롭 경계선
        crop_pen = QPen(QColor(100, 200, 100), 2.5, Qt.PenStyle.DashLine)
        crop_rect = self.scene.addRect(0, 0, w, h, crop_pen)
        crop_rect.setZValue(10)
        
        # 2. 중심 파란 십자선
        center_pen = QPen(QColor(70, 100, 240, 160), 1.5, Qt.PenStyle.DashLine)
        self.scene.addLine(w / 2, -15, w / 2, h + 15, center_pen).setZValue(10)
        self.scene.addLine(-25, h * 0.45, w + 25, h * 0.45, center_pen).setZValue(10)
        
        # 3. 유형별 국가 표준 비율 기하학적 연산 변환
        is_passport = "여권" in self.current_preset_name or "운전" in self.current_preset_name or "주민" in self.current_preset_name
        is_visa = "비자" in self.current_preset_name
        is_pet = "증멍증냥" in self.current_preset_name
        
        if is_passport:
            face_w, face_h = w * 0.60, h * 0.58
            face_y = h * 0.18
        elif is_visa:
            face_w, face_h = w * 0.55, h * 0.55
            face_y = h * 0.15
        elif is_pet:
            face_w, face_h = w * 0.52, h * 0.52
            face_y = h * 0.28
        else:
            face_w, face_h = w * 0.42, h * 0.40
            face_y = h * 0.22
            
        face_x = (w - face_w) / 2
        top_red_y = face_y
        bottom_red_y = face_y + face_h
        
        # 빨간선 생성
        red_pen = QPen(QColor(230, 70, 70, 200), 1.5, Qt.PenStyle.DashLine)
        self.scene.addLine(-25, top_red_y, w + 25, top_red_y, red_pen).setZValue(10)
        self.scene.addLine(-25, bottom_red_y, w + 25, bottom_red_y, red_pen).setZValue(10)
        
        # 4. 안면 주황색 윤곽선 (점선)
        face_contour_pen = QPen(QColor(255, 165, 0, 180), 1.8, Qt.PenStyle.DashLine)
        ellipse_guide = self.scene.addEllipse(face_x, face_y, face_w, face_h, face_contour_pen)
        ellipse_guide.setZValue(11)
        

        # 어깨선 가이드
        shoulder_path = QPainterPath()
        if is_passport or is_visa:
            shoulder_path.moveTo(face_x - 30, h)
            shoulder_path.quadTo(face_x + 10, face_y + face_h - 10, face_x + 15, face_y + face_h)
            shoulder_path.moveTo(face_x + face_w + 30, h)
            shoulder_path.quadTo(face_x + face_w - 10, face_y + face_h - 10, face_x + face_w - 15, face_y + face_h)
        elif is_pet:
            shoulder_path.moveTo(face_x - 10, h)
            shoulder_path.quadTo(w/2, face_y + face_h + 10, face_x + face_w + 10, h)
        else:
            shoulder_path.moveTo(0, h * 0.90)
            shoulder_path.quadTo(face_x - 10, face_y + face_h + 10, face_x + 10, face_y + face_h)
            shoulder_path.moveTo(w, h * 0.90)
            shoulder_path.quadTo(face_x + face_w + 10, face_y + face_h + 10, face_x + face_w - 10, face_y + face_h)
        self.scene.addPath(shoulder_path, face_contour_pen).setZValue(11)

        # 가이드라인 설명 글자 크기를 12pt로 고정
        font_size = 10
        y_offset = int(font_size * 1.5)
        x_offset = int(font_size * 8.3)

        def create_guide_text(text, x, y, color_hex, is_right_side=False):
            txt_item = QGraphicsTextItem(text)
            txt_item.setFont(QFont("Malgun Gothic", font_size, QFont.Weight.Bold))
            txt_item.setDefaultTextColor(QColor(color_hex))
            if is_right_side:
                txt_item.setPos(x + 10, y - y_offset)
            else:
                txt_item.setPos(x - x_offset, y - y_offset)
            txt_item.setZValue(12)
            self.scene.addItem(txt_item)

        create_guide_text("중심 수평선 ▶", -25, h * 0.45, "#4664f0", is_right_side=False)
        create_guide_text("◀ 정수리 기준선", w + 25, top_red_y, "#e64646", is_right_side=True)
        create_guide_text("◀ 턱끝 기준선", w + 25, bottom_red_y, "#e64646", is_right_side=True)
        
        # '▼ 좌우 대칭 중심선' 한 줄로 표시, ▼를 중심선(w/2)에 정렬
        combined_item = QGraphicsTextItem("▼ 좌우 대칭 중심선")
        combined_item.setFont(QFont("Malgun Gothic", font_size, QFont.Weight.Bold))
        combined_item.setDefaultTextColor(QColor("#4664f0"))
        combined_rect = combined_item.boundingRect()

        # ▼ 문자 하나의 너비를 계산하기 위한 임시 아이템
        arrow_tmp = QGraphicsTextItem("▼")
        arrow_tmp.setFont(QFont("Malgun Gothic", font_size, QFont.Weight.Bold))
        arrow_w = arrow_tmp.boundingRect().width()

        # ▼의 중심이 중심선(w/2)에 오도록 x 좌표 계산
        combined_x = (w / 2) - (arrow_w / 2)

        padding = max(6, int(font_size * 0.6))
        # 텍스트가 초록 크롭 경계(상단 y=0) 바깥(위쪽)에 위치하도록 설정
        combined_y = -(combined_rect.height() + padding)

        combined_item.setPos(combined_x, combined_y)
        combined_item.setZValue(12)
        self.scene.addItem(combined_item)

        self.fit_guide_in_view()
        self.toggle_guides(self.chk_guide.checkState().value)

    def fit_guide_in_view(self):
        if hasattr(self, 'current_target_size') and self.view:
            w, h = self.current_target_size
            self.view.fitInView(QRectF(-140, -50, w + 280, h + 100), Qt.AspectRatioMode.KeepAspectRatio)
            self.view.centerOn(w / 2, h / 2)

    def open_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "사진 불러오기", "", "이미지 파일 (*.jpg *.jpeg *.png *.bmp)")
        if file_path:
            pixmap = QPixmap(file_path)
            if not pixmap.isNull():
                self.view.set_image(pixmap)
                self.update_guide_lines()

    def save_image(self):
        if not self.view.image_item:
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "사진 저장하기", f"result_{self.current_preset_name.split(' ')[0]}.jpg", "이미지 파일 (*.jpg *.png)")
        if file_path:
            w, h = self.current_target_size
            output_image = QImage(w, h, QImage.Format.Format_ARGB32)
            output_image.fill(Qt.GlobalColor.transparent)
            
            painter = QPainter(output_image)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
            
            for item in self.scene.items():
                if isinstance(item, (QGraphicsLineItem, QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem, QGraphicsTextItem)):
                    item.setVisible(False)
            
            self.scene.render(painter, QRectF(0, 0, w, h), QRectF(0, 0, w, h))
            self.toggle_guides(self.chk_guide.checkState().value)
            painter.end()
            output_image.save(file_path)

    def apply_image_style_theme(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #1a1c1e; }
            QMenuBar { background-color: #111215; color: #b3b4b6; font-family: 'Malgun Gothic'; font-size: 14px; border-bottom: 1px solid #222327; }
            QMenuBar::item { padding: 8px 14px; }
            QMenuBar::item:selected { background-color: #222327; color: #ffffff; }
            
            QLabel { color: #a5a7aa; font-family: 'Malgun Gothic'; font-size: 14px; font-weight: bold; }
            QComboBox { background-color: #2c2f36; color: #ffffff; border: 1px solid #3e424b; border-radius: 4px; padding: 12px; font-size: 15px; }
            QPushButton#ShortcutBtn { background-color: #2c2f36; color: #a5a7aa; border: 1px solid #3e424b; border-radius: 4px; padding: 12px; font-size: 15px; }
            QPushButton#ShortcutBtn:hover { background-color: #383c45; color: #ffffff; }
            
            QPushButton#ActionBlueBtn {
                background-color: #1a73e8; color: #ffffff; border: none; border-radius: 4px; padding: 15px; font-size: 15px; font-weight: bold;
            }
            QPushButton#ActionBlueBtn:hover { background-color: #297be6; }
            QCheckBox { color: #a5a7aa; font-size: 15px; spacing: 10px; }
            QSlider::groove:horizontal { height: 8px; background: #2c2f36; border-radius: 4px; }
            QSlider::handle:horizontal { background: #1a73e8; width: 22px; margin-top: -7px; margin-bottom: -7px; border-radius: 11px; }
            
            QScrollArea#RightScrollArea { border: none; background-color: #151619; }
            #RightPanel { background-color: #151619; }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PhotoProcessorApp()
    window.show()
    sys.exit(app.exec())