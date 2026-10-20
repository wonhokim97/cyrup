import os
import json
from datetime import datetime
from importlib import import_module
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageOps, ImageDraw, ImageChops, ImageFilter

class IcoConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("고화질 프리미엄 Icon 변환기")
        self.root.geometry("340x380")
        self.root.resizable(False, False)
        
        # 윈도우 종료 시 자동으로 설정을 저장하도록 이벤트 바인딩
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 이미지 데이터 변수
        self.original_img = None
        self.square_original_img = None
        self.preview_base_img = None
        self.preview_tk = None
        self.file_path = None

        # ---------------- 0. 설정 로드 ----------------
        self.load_settings()

        # ---------------- 1. 색상 프리셋 데이터 정의 ----------------
        self.bg_presets = {
            "original": {"name": "투명 배경 (원본 - 기본)", "color": None},
            "white": {"name": "흰색 (White)", "color": "#FFFFFF"},
            "black": {"name": "검은색 (Black)", "color": "#1A1A1A"},
            "red": {"name": "빨간색 (Red)", "color": "#E53935"},
            "orange": {"name": "주황색 (Orange)", "color": "#FB8C00"},
            "yellow": {"name": "노란색 (Yellow)", "color": "#FDD835"},
            "green": {"name": "초록색 (Green)", "color": "#43A047"},
            "blue": {"name": "파란색 (Blue)", "color": "#1E88E5"},
            "purple": {"name": "보라색 (Purple)", "color": "#8E24AA"},
            "gray": {"name": "회색 (Gray)", "color": "#757575"},
            "custom": {"name": "사용자 지정 색상", "color": self.settings["custom_bg_color"]}
        }

        self.icon_presets = {
            "original": {"name": "기본 색상 (원본)", "color": None},
            "white": {"name": "흰색 (White)", "color": "#FFFFFF"},
            "black": {"name": "검은색 (Black)", "color": "#1A1A1A"},
            "red": {"name": "빨간색 (Red)", "color": "#E53935"},
            "orange": {"name": "주황색 (Orange)", "color": "#FB8C00"},
            "yellow": {"name": "노란색 (Yellow)", "color": "#FDD835"},
            "green": {"name": "초록색 (Green)", "color": "#43A047"},
            "blue": {"name": "파란색 (Blue)", "color": "#1E88E5"},
            "purple": {"name": "보라색 (Purple)", "color": "#8E24AA"},
            "grayscale": {"name": "흑백 (Grayscale)", "color": "gray"},
            "custom": {"name": "사용자 지정 색상", "color": self.settings["custom_icon_color"]}
        }

        # 메뉴용 실시간 고화질 컬러 아이콘 이미지 생성
        self.bg_icons = {}
        for key, val in self.bg_presets.items():
            self.bg_icons[key] = self.make_color_icon(val["color"])
            
        self.icon_icons = {}
        for key, val in self.icon_presets.items():
            self.icon_icons[key] = self.make_color_icon(val["color"])

        # ---------------- 2. 메뉴바 및 풀다운 메뉴 구성 ----------------
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        # [파일] 메뉴 (PNG, JPG, WEBP, GIF, BMP, ICO 지원)
        self.file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일(F)", menu=self.file_menu)
        self.file_menu.add_command(label="파일 열기...", command=self.select_file)
        self.file_menu.add_command(label="ICO로 저장...", command=self.convert_and_save, state="disabled")
        self.file_menu.add_separator()
        self.file_menu.add_command(label="종료(X)", command=self.on_close)

        # [아이콘 크기] 메뉴
        self.size_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="크기(S)", menu=self.size_menu)
        self.size_options = [16, 32, 48, 64, 128, 256, 512]
        self.size_vars = {}
        for size in self.size_options:
            is_checked = size in self.settings["sizes"]
            var = tk.BooleanVar(value=is_checked)
            self.size_vars[size] = var
            self.size_menu.add_checkbutton(label=f"{size} x {size}", variable=var)

        # [모양 설정] 메뉴 (자동 꽉 채우기 및 선 크롭 기술 적용)
        self.shape_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="모양(M)", menu=self.shape_menu)
        self.shape_var = tk.StringVar(value=self.settings["shape"])
        self.shape_menu.add_radiobutton(label="기본 모양 (원본 비율)", variable=self.shape_var, value="default", command=self.update_preview)
        self.shape_menu.add_radiobutton(label="아이폰 카드 - 단색 채우기 (여백 없이 최대화) ★", variable=self.shape_var, value="card_solid", command=self.update_preview)
        self.shape_menu.add_radiobutton(label="아이폰 카드 - 투명 글래스 (여백 없이 최대화) ★", variable=self.shape_var, value="card_glass", command=self.update_preview)
        
        self.shape_menu.add_separator()
        
        self.iphone_round_var = tk.BooleanVar(value=self.settings.get("iphone_round", True))
        self.shape_menu.add_checkbutton(
            label="아이폰 스타일 (자동 둥근 모서리 적용) ★", 
            variable=self.iphone_round_var, 
            command=self.update_preview
        )

        # ---------------- [통합] 색상 설정 대메뉴 ----------------
        self.color_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="색상(C)", menu=self.color_menu)

        # ① 자동 바탕 제거 토글 버튼
        self.bg_remove_var = tk.BooleanVar(value=self.settings["bg_remove"])
        self.color_menu.add_checkbutton(
            label="자동 바탕 제거 실행 (On/Off) ★", 
            variable=self.bg_remove_var, 
            command=self.update_preview
        )
        self.color_menu.add_separator()

        # ② 바탕색 프리셋 및 사용자 지정 정의
        self.bg_submenu = tk.Menu(self.color_menu, tearoff=0)
        self.color_menu.add_cascade(label="바탕색 변경", menu=self.bg_submenu)
        self.bg_var = tk.StringVar(value=self.settings["bg_color"])
        
        for key, val in self.bg_presets.items():
            if key == "custom":
                continue
            self.bg_submenu.add_radiobutton(
                label="  " + val["name"], 
                image=self.bg_icons[key],
                compound="left",
                variable=self.bg_var, 
                value=key,
                command=self.update_preview
            )
        self.bg_submenu.add_separator()
        self.bg_submenu.add_radiobutton(
            label="  사용자 지정 색상 선택... ", 
            image=self.bg_icons["custom"],
            compound="left",
            variable=self.bg_var, 
            value="custom",
            command=self.choose_custom_bg
        )

        self.color_menu.add_separator()

        # ③ 아이콘 색상 프리셋 및 사용자 지정 정의
        self.icon_submenu = tk.Menu(self.color_menu, tearoff=0)
        self.color_menu.add_cascade(label="아이콘 색상 변경", menu=self.icon_submenu)
        self.icon_var = tk.StringVar(value=self.settings["icon_color"])
        
        for key, val in self.icon_presets.items():
            if key == "custom":
                continue
            self.icon_submenu.add_radiobutton(
                label="  " + val["name"], 
                image=self.icon_icons[key],
                compound="left",
                variable=self.icon_var, 
                value=key,
                command=self.update_preview
            )
        self.icon_submenu.add_separator()
        self.icon_submenu.add_radiobutton(
            label="  사용자 지정 색상 선택... ", 
            image=self.icon_icons["custom"],
            compound="left",
            variable=self.icon_var, 
            value="custom",
            command=self.choose_custom_icon
        )

        # ---------------- 3. 화면 중앙 레이아웃 ----------------
        info_label = tk.Label(
            root, 
            text="상단 메뉴바에서 파일 열기, 모양, 색상을 설정하세요.", 
            font=("Malgun Gothic", 8), 
            fg="gray", 
            pady=8
        )
        info_label.pack()

        # 256x256 격자무늬 캔버스
        self.preview_canvas = tk.Canvas(
            root, width=256, height=256, bg="#e0e0e0", 
            highlightthickness=1, highlightbackground="#cccccc"
        )
        self.preview_canvas.pack(pady=5)
        self.draw_checkerboard()

        # 파일 정보 라벨
        self.file_label = tk.Label(
            root, 
            text="선택된 파일 없음", 
            font=("Malgun Gothic", 9, "bold"), 
            fg="gray", 
            wraplength=300, 
            pady=8
        )
        self.file_label.pack()

        # ---------------- 4. 맨 아래 한줄 로그 상태바 ----------------
        self.status_label = tk.Label(
            root, 
            text="준비 완료", 
            font=("Malgun Gothic", 8), 
            fg="#555555", 
            bg="#f0f0f0", 
            anchor="w", 
            padx=10, 
            pady=3,
            bd=1,
            relief="sunken"
        )
        self.status_label.pack(side="bottom", fill="x")

        self.log("프로그램이 실행되었습니다. 이전 설정을 불러왔습니다.")

    # ---------------- 정밀 컬러 메뉴 아이콘 빌더 ----------------
    def make_color_icon(self, color_hex, size=14):
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        if color_hex is None:
            draw.ellipse([1, 1, size-2, size-2], outline=(100, 100, 100, 255), width=2)
            draw.ellipse([size//2-1, size//2-1, size//2+1, size//2+1], fill=(100, 100, 100, 255))
        elif color_hex == "gray":
            draw.ellipse([0, 0, size-1, size-1], fill=(128, 128, 128, 255))
            draw.ellipse([0, 0, size-1, size-1], outline=(80, 80, 80, 255), width=1)
        else:
            h = color_hex.lstrip('#')
            rgb = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
            draw.ellipse([0, 0, size-1, size-1], fill=rgb + (255,))
            
            brightness = 0.299*rgb[0] + 0.587*rgb[1] + 0.114*rgb[2]
            border_color = (180, 180, 180, 255) if brightness > 220 else (240, 240, 240, 50)
            draw.ellipse([0, 0, size-1, size-1], outline=border_color, width=1)
            
        return ImageTk.PhotoImage(img)

    # ---------------- 로그 출력용 헬퍼 함수 ----------------
    def log(self, message, log_type="info"):
        if log_type == "error":
            self.status_label.config(text=f"❌ {message}", fg="#D32F2F")
        elif log_type == "success":
            self.status_label.config(text=f"✔ {message}", fg="#2E7D32")
        elif log_type == "action":
            self.status_label.config(text=f"⚡ {message}", fg="#0078D7")
        else:
            self.status_label.config(text=message, fg="#555555")

    # ---------------- 사용자 지정 색상 선택 및 실시간 메뉴 아이콘 재생성 ----------------
    def choose_custom_bg(self):
        color = colorchooser.askcolor(
            title="바탕색 사용자 지정 선택", 
            initialcolor=self.settings["custom_bg_color"]
        )
        if color[1]:
            self.settings["custom_bg_color"] = color[1]
            self.bg_presets["custom"]["color"] = color[1]
            self.bg_var.set("custom")
            
            self.bg_icons["custom"] = self.make_color_icon(color[1])
            self.bg_submenu.entryconfigure("end", image=self.bg_icons["custom"])
            
            self.log(f"바탕색이 사용자 지정 색상({color[1]})으로 변경되었습니다.", "action")
        self.update_preview()

    def choose_custom_icon(self):
        color = colorchooser.askcolor(
            title="아이콘 색상 사용자 지정 선택", 
            initialcolor=self.settings["custom_icon_color"]
        )
        if color[1]:
            self.settings["custom_icon_color"] = color[1]
            self.icon_presets["custom"]["color"] = color[1]
            self.icon_var.set("custom")
            
            self.icon_icons["custom"] = self.make_color_icon(color[1])
            self.icon_submenu.entryconfigure("end", image=self.icon_icons["custom"])
            
            self.log(f"아이콘 색상이 사용자 지정 색상({color[1]})으로 변경되었습니다.", "action")
        self.update_preview()

    # ---------------- 설정 파일 입출력 로직 ----------------
    def get_config_path(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            return os.path.join(script_dir, "ico_converter_config.json")
        except NameError:
            return "ico_converter_config.json"

    def load_settings(self):
        config_path = self.get_config_path()
        self.defaults = {
            "sizes": [16, 32, 48, 256],
            "shape": "default",
            "iphone_round": True,
            "bg_color": "original",
            "icon_color": "original",
            "bg_remove": False,
            "custom_bg_color": "#FFFFFF",
            "custom_icon_color": "#0078D7"
        }
        
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    self.settings = json.load(f)
                    
                    if "color" in self.settings:
                        self.settings["icon_color"] = self.settings.pop("color")
                    
                    if isinstance(self.settings.get("bg_remove"), str):
                        self.settings["bg_remove"] = self.settings["bg_remove"] != "none"
                    
                    if self.settings.get("shape") in ["iphone", "clip"]:
                        self.settings["shape"] = "default"
                    elif self.settings.get("shape") not in ["default", "card_solid", "card_glass"]:
                        self.settings["shape"] = "default"
                        
                    for k, v in self.defaults.items():
                        if k not in self.settings:
                            self.settings[k] = v
            except Exception:
                self.settings = self.defaults.copy()
        else:
            self.settings = self.defaults.copy()

    def save_settings(self):
        config_path = self.get_config_path()
        current_sizes = [size for size, var in self.size_vars.items() if var.get()]
        
        current_settings = {
            "sizes": current_sizes,
            "shape": self.shape_var.get(),
            "bg_color": self.bg_var.get(),
            "icon_color": self.icon_var.get(),
            "bg_remove": self.bg_remove_var.get(),
            "iphone_round": self.iphone_round_var.get(),
            "custom_bg_color": self.settings["custom_bg_color"],
            "custom_icon_color": self.settings["custom_icon_color"]
        }
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(current_settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"설정 저장 실패: {e}")

    def on_close(self):
        self.save_settings()
        self.root.destroy()

    # ---------------- 그래픽 가공 및 제어 로직 ----------------
    def draw_checkerboard(self):
        self.preview_canvas.delete("all")
        tile_size = 16
        for i in range(0, 256, tile_size):
            for j in range(0, 256, tile_size):
                if (i // tile_size + j // tile_size) % 2 == 0:
                    self.preview_canvas.create_rectangle(i, j, i+tile_size, j+tile_size, fill="#f9f9f9", outline="")
                else:
                    self.preview_canvas.create_rectangle(i, j, i+tile_size, j+tile_size, fill="#e5e5e5", outline="")

    def make_square_with_transparency(self, img):
        width, height = img.size
        max_dim = max(width, height)
        
        square_img = Image.new("RGBA", (max_dim, max_dim), (0, 0, 0, 0))
        offset_x = (max_dim - width) // 2
        offset_y = (max_dim - height) // 2
        square_img.paste(img, (offset_x, offset_y), img)
        
        return square_img

    def crop_transparent_borders(self, img, threshold=15):
        img = img.convert("RGBA")
        r, g, b, a = img.split()
        mask = a.point(lambda p: 255 if p >= threshold else 0)
        
        bbox = mask.getbbox()
        if bbox:
            return img.crop(bbox)
        return img

    def apply_bg_removal_auto(self, img):
        if not self.bg_remove_var.get():
            return img
            
        try:
            # 선택적 AI 패키지: 미설치 시 기존 기본 제거 방식 사용
            remove = import_module("rembg").remove
            self.log("AI 엔진 가동 중... 정밀하게 배경을 분석해 잘라냅니다.", "action")
            self.root.update_idletasks()
            img_rgba = img.convert("RGBA")
            result = remove(img_rgba)
            self.log("AI 정밀 바탕 제거 성공!", "success")
            return result
        except ImportError:
            self.log("경계 보존형 마술봉 엔진으로 배경 제거 중...", "action")
            
        img = img.convert("RGBA")
        w, h = img.size
        
        tolerance = 28
        filled_img = img.copy()
        
        seeds = [
            (0, 0), (w - 1, 0), 
            (0, h - 1), (w - 1, h - 1),
            (2, 2), (w - 3, 2), 
            (2, h - 3), (w - 3, h - 3)
        ]
        
        for seed in seeds:
            if 0 <= seed[0] < w and 0 <= seed[1] < h:
                if filled_img.getpixel(seed)[3] < 50:
                    continue
                ImageDraw.floodfill(filled_img, xy=seed, value=(0, 0, 0, 0), thresh=tolerance)
                
        r, g, b, a = filled_img.split()
        smoothed_a = a.filter(ImageFilter.MedianFilter(size=3))
        smoothed_a = smoothed_a.filter(ImageFilter.GaussianBlur(radius=0.6))
        
        clean_img = Image.merge("RGBA", (r, g, b, smoothed_a))
        
        self.log("바탕 제거 및 가장자리 부드럽게 정화 완료!", "success")
        return clean_img

    def apply_clip_shape(self, img):
        width, height = img.size
        scale = 4
        mask_size = (width * scale, height * scale)
        
        mask = Image.new("L", mask_size, 0)
        draw = ImageDraw.Draw(mask)
        radius = int(mask_size[0] * 0.22)
        draw.rounded_rectangle((0, 0, mask_size[0], mask_size[1]), radius=radius, fill=255)
        
        mask = mask.resize((width, height), Image.Resampling.LANCZOS)
        
        r, g, b, a = img.split()
        new_a = ImageChops.darker(a, mask)
        
        return Image.merge("RGBA", (r, g, b, new_a))

    def apply_iphone_shape_smart(self, img, is_silent=False):
        img = img.convert("RGBA")
        w, h = img.size
        
        corners_alpha = [
            img.getpixel((0, 0))[3],
            img.getpixel((w - 1, 0))[3],
            img.getpixel((0, h - 1))[3],
            img.getpixel((w - 1, h - 1))[3]
        ]
        
        if any(alpha < 100 for alpha in corners_alpha):
            if not is_silent:
                self.log("배경이 없는 이미지입니다. 모서리 깎기를 생략(원형 보존)합니다.", "info")
            return img
        else:
            if not is_silent:
                self.log("배경이 감지되어 모서리를 아이폰 스타일로 둥글게 깎았습니다.", "action")
            return self.apply_clip_shape(img)

    def apply_card_shape(self, img, bg_type, bg_color, icon_color):
        cropped_logo = self.crop_transparent_borders(img)
        logo_w, logo_h = cropped_logo.size
        
        max_dim = max(logo_w, logo_h)
        scale = 4
        card_dim = max_dim * scale
        
        if bg_type == "glass":
            card_bg = Image.new("RGBA", (card_dim, card_dim), (255, 255, 255, 60))
        else:
            if bg_color:
                h = bg_color.lstrip('#')
                rgb = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
                card_bg = Image.new("RGBA", (card_dim, card_dim), rgb + (255,))
            else:
                card_bg = Image.new("RGBA", (card_dim, card_dim), (0, 0, 0, 0))
                
        mask = Image.new("L", (card_dim, card_dim), 0)
        draw_mask = ImageDraw.Draw(mask)
        radius = int(card_dim * 0.22)
        draw_mask.rounded_rectangle((0, 0, card_dim, card_dim), radius=radius, fill=255)
        
        if bg_type == "glass":
            draw_bg = ImageDraw.Draw(card_bg)
            draw_bg.rounded_rectangle(
                (0, 0, card_dim, card_dim), 
                radius=radius, 
                outline=(255, 255, 255, 180), 
                width=max(4, scale)
            )
            
        if icon_color is not None:
            cropped_logo = self.apply_tint(cropped_logo, icon_color)
            
        max_logo_dim = int(card_dim * 0.88)
        ratio = min(max_logo_dim / logo_w, max_logo_dim / logo_h)
        final_logo_w = int(logo_w * ratio)
        final_logo_h = int(logo_h * ratio)
        resized_logo = cropped_logo.resize((final_logo_w, final_logo_h), Image.Resampling.LANCZOS)
        
        offset_x = (card_dim - final_logo_w) // 2
        offset_y = (card_dim - final_logo_h) // 2
        card_bg.paste(resized_logo, (offset_x, offset_y), resized_logo)
        
        final_card = card_bg.resize((max_dim, max_dim), Image.Resampling.LANCZOS)
        final_mask = mask.resize((max_dim, max_dim), Image.Resampling.LANCZOS)
        
        r, g, b, a = final_card.split()
        new_a = ImageChops.darker(a, final_mask)
        
        return Image.merge("RGBA", (r, g, b, new_a))

    def apply_tint(self, img, target_color):
        if target_color is None:
            return img

        r, g, b, a = img.split()
        rgb = Image.merge("RGB", (r, g, b))
        gray = rgb.convert("L")

        if target_color == "gray":
            colorized = rgb.convert("L").convert("RGB")
        else:
            colorized = ImageOps.colorize(gray, black="black", white=target_color)

        r_n, g_n, b_n = colorized.split()
        return Image.merge("RGBA", (r_n, g_n, b_n, a))

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="이미지 파일(PNG, JPG, WEBP, GIF, BMP, ICO)을 선택하세요",
            filetypes=[
                ("지원하는 모든 파일", "*.png;*.jpg;*.jpeg;*.webp;*.gif;*.bmp;*.ico"),
                ("PNG 파일", "*.png"),
                ("JPG/JPEG 파일", "*.jpg;*.jpeg"),
                ("WEBP 파일", "*.webp"),
                ("GIF 파일", "*.gif"),
                ("BMP 파일", "*.bmp"),
                ("ICO 파일", "*.ico")
            ]
        )

        if not file_path:
            self.log("파일 선택이 취소되었습니다.", "info")
            return

        self.file_path = file_path
        file_name = os.path.basename(file_path)
        self.file_label.config(text=f"선택됨: {file_name}", fg="#0078d7")

        try:
            raw_img = Image.open(file_path)
            
            # ICO 형식 다중 프레임 대응
            frame_count = int(getattr(raw_img, "n_frames", 1))
            if frame_count > 1:
                largest_img = None
                max_size = 0
                for frame in range(frame_count):
                    raw_img.seek(frame)
                    w, h = raw_img.size
                    if w > max_size:
                        max_size = w
                        largest_img = raw_img.copy()
                self.original_img = largest_img.convert("RGBA") if largest_img else raw_img.convert("RGBA")
            else:
                self.original_img = raw_img.convert("RGBA")

            w, h = self.original_img.size
            ratio = min(256 / w, 256 / h)
            new_w, new_h = int(w * ratio), int(h * ratio)
            self.preview_base_img = self.original_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            
            self.update_preview()
            self.file_menu.entryconfig("ICO로 저장...", state="normal")
            self.log(f"성공적으로 불러왔습니다: {file_name}", "success")

        except Exception as e:
            self.log(f"파일을 여는 중 에러 발생: {e}", "error")
            messagebox.showerror("오류", f"파일을 불러오는 중 오류가 발생했습니다:\n{e}")

    def process_image(self, img, shape, bg_color, icon_color, is_bg_remove, is_iphone_round):
        # 1단계: 자동 바탕색 투명화 연산 처리
        processed = self.apply_bg_removal_auto(img) if is_bg_remove else img

        # 2단계: 아이폰 스타일 (자동 둥근 모서리) 연산 선행 처리
        if is_iphone_round:
            processed = self.apply_iphone_shape_smart(processed, is_silent=(shape != "default"))

        # 3단계: 모양 선택에 따른 최종 렌더링 파이프라인
        if shape == "default":
            processed_square = self.make_square_with_transparency(processed)
            return self.apply_tint(processed_square, icon_color)
            
        elif shape == "card_solid":
            return self.apply_card_shape(processed, "solid", bg_color, icon_color)
            
        elif shape == "card_glass":
            return self.apply_card_shape(processed, "glass", bg_color, icon_color)
            
        return self.make_square_with_transparency(processed)

    def update_preview(self):
        if self.preview_base_img is None:
            return

        shape = self.shape_var.get()
        is_bg_remove = self.bg_remove_var.get()
        is_iphone_round = self.iphone_round_var.get()
        
        bg_key = self.bg_var.get()
        bg_color = self.bg_presets[bg_key]["color"]
        
        icon_key = self.icon_var.get()
        icon_color = self.icon_presets[icon_key]["color"]

        modified_preview = self.process_image(self.preview_base_img, shape, bg_color, icon_color, is_bg_remove, is_iphone_round)
        
        self.preview_tk = ImageTk.PhotoImage(modified_preview)
        
        self.draw_checkerboard()
        self.preview_canvas.create_image(128, 128, image=self.preview_tk)
        self.root.update_idletasks()

    def convert_and_save(self):
        file_path = self.file_path
        original_img = self.original_img
        if original_img is None or not file_path:
            return

        selected_sizes = [size for size, var in self.size_vars.items() if var.get()]
        if not selected_sizes:
            self.log("저장 실패: 최소 한 개 이상의 크기를 선택해야 합니다.", "error")
            messagebox.showwarning("경고", "최소 하나 이상의 아이콘 크기를 선택해야 합니다.")
            return

        selected_sizes.sort()

        try:
            folder, filename = os.path.split(file_path)
            name, _ = os.path.splitext(filename)

            # [업데이트] 저장 시 파일 이름에 모양 정보 및 현재 날짜 자동 연동 (영문 적용)
            today_str = datetime.now().strftime("%Y%m%d")
            shape_key = self.shape_var.get()
            shape_names = {
                "default": "default",
                "card_solid": "solid",
                "card_glass": "glass"
            }
            shape_info = shape_names.get(shape_key, "default")
            
            # 아이폰 둥근 모서리 옵션이 켜져 있고 기본 모양일 때 파일명 보정
            if self.iphone_round_var.get() and shape_key == "default":
                shape_info = "round"

            # 최종 자동 제안 파일명 조합 (예: logo_glass_20260608.ico)
            suggested_filename = f"{name}_{shape_info}_{today_str}.ico"

            save_path = filedialog.asksaveasfilename(
                title="ICO / PNG 파일 저장 위치를 선택하세요",
                initialdir=folder,
                initialfile=suggested_filename,
                filetypes=[("ICO 파일", "*.ico")]
            )

            if not save_path:
                self.log("아이콘 저장이 취소되었습니다.", "info")
                return

            shape = self.shape_var.get()
            is_bg_remove = self.bg_remove_var.get()
            is_iphone_round = self.iphone_round_var.get()
            
            bg_key = self.bg_var.get()
            bg_color = self.bg_presets[bg_key]["color"]
            
            icon_key = self.icon_var.get()
            icon_color = self.icon_presets[icon_key]["color"]

            modified_original = self.process_image(original_img, shape, bg_color, icon_color, is_bg_remove, is_iphone_round)

            resized_images = []
            for size in selected_sizes:
                resized = modified_original.resize((size, size), Image.Resampling.LANCZOS)
                resized_images.append(resized)

            resized_images[-1].save(
                save_path, 
                format="ICO", 
                sizes=[(s, s) for s in selected_sizes], 
                append_images=resized_images[:-1]
            )

            self.save_settings()
            self.log(f"아이콘 저장 완료! ({os.path.basename(save_path)})", "success")
            messagebox.showinfo("성공", f"아이콘 저장이 완료되었습니다!\n\n저장 경로:\n{save_path}")

        except Exception as e:
            self.log(f"아이콘 변환 중 에러 발생: {e}", "error")
            messagebox.showerror("오류", f"변환 중 오류가 발생했습니다:\n{e}")

if __name__ == "__main__":
    win = tk.Tk()
    app = IcoConverterApp(win)
    win.mainloop()
