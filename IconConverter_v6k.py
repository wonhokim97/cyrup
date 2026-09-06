import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageOps, ImageOps as PILImageOps
import os
import io
from datetime import datetime

class IconGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional Icon Generator")
        
        # 기본 기준 사이즈 설정
        self.base_width = 430
        self.base_height = 430
        self.root.geometry(f"{self.base_width}x{self.base_height}")
        self.root.resizable(False, False)

        # 초기 변수 설정
        self.file_path = None
        self.original_img = None  
        self.source_img = None    
        self.current_img = None   
        
        self.bg_color = None
        self.icon_color = None
        self.is_transparent_mode = False 
        self.icon_size = tk.IntVar(value=256)
        self.style_var = tk.StringVar(value="Normal")
        
        self.font_size = 11
        
        # 메뉴 아이콘 이미지 가비지 컬렉션 방지용 리스트
        self.menu_icons_refs = []

        self.setup_menubar()
        
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(expand=True)
        
        self.preview_frame = tk.Frame(self.main_container, width=300, height=300, bg="#dcdcdc", highlightbackground="gray", highlightthickness=1)
        self.preview_frame.pack(pady=(5, 10))
        self.preview_frame.pack_propagate(False)
        
        self.lbl_preview = tk.Label(self.preview_frame, text="이미지를 불러와주세요", fg="#555555")
        self.lbl_preview.pack(expand=True, fill=tk.BOTH)
        
        self.fixed_bg_image = self.create_checkerboard((296, 296))
        self.set_default_checkerboard()

        self.btn_frame = tk.Frame(self.main_container)
        self.btn_frame.pack(pady=(5, 5))

        self.btn_reset_bg = tk.Button(self.btn_frame, text="🎨 배경색 초기화", bg="#f5f5f5", padx=8, pady=5, relief=tk.GROOVE, command=self.reset_bg_color)
        self.btn_reset_bg.pack(side=tk.LEFT, padx=5)

        self.btn_reset_all = tk.Button(self.btn_frame, text="🔄 원본 복귀", bg="#e1e1e1", padx=12, pady=5, relief=tk.GROOVE, command=self.reset_to_original)
        self.btn_reset_all.pack(side=tk.LEFT, padx=5)

        self.btn_reset_icon = tk.Button(self.btn_frame, text="🎯 아이콘색 초기화", bg="#f5f5f5", padx=8, pady=5, relief=tk.GROOVE, command=self.reset_icon_color)
        self.btn_reset_icon.pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar(value="준비됨")
        self.status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.setup_ui_styles()

    def create_menu_color_icon(self, color_tuple, special=None):
        size = int(self.font_size * 1.2)
        if special == "transparent":
            img = self.create_checkerboard((size, size), grid_size=size//2)
        elif special == "rainbow":
            img = Image.new('RGBA', (size, size), (0,0,0,0))
            draw = ImageDraw.Draw(img)
            draw.rectangle([0, 0, size//2, size//2], fill=(255,0,0,255))
            draw.rectangle([size//2, 0, size, size//2], fill=(0,255,0,255))
            draw.rectangle([0, size//2, size//2, size], fill=(0,0,255,255))
            draw.rectangle([size//2, size//2, size, size], fill=(255,255,0,255))
        else:
            img = Image.new('RGBA', (size, size), color_tuple + (255,))
            ImageDraw.Draw(img).rectangle([0, 0, size-1, size-1], outline=(150,150,150,255))

        tk_img = ImageTk.PhotoImage(img)
        self.menu_icons_refs.append(tk_img)
        return tk_img

    def setup_menubar(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        dropdown_font = ("맑은 고딕", self.font_size)
        
        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0, font=dropdown_font)
        file_menu.add_command(label=" 파일 불러오기", command=self.load_file)
        file_menu.add_command(label=" 다른 이름으로 저장", command=self.save_file)
        file_menu.add_separator()
        file_menu.add_command(label=" 프로그램 종료", command=self.root.quit)
        menubar.add_cascade(label="파일", menu=file_menu)
        
        # 변환 메뉴
        opt_menu = tk.Menu(menubar, tearoff=0, font=dropdown_font)
        
        # 아이콘 크기 선택
        res_menu = tk.Menu(opt_menu, tearoff=0, font=dropdown_font)
        for s in [64, 128, 256, 512]:
            res_menu.add_radiobutton(label=f"{s}x{s} px", variable=self.icon_size, value=s)
        opt_menu.add_cascade(label=" 아이콘 크기 선택", menu=res_menu)
        
        # 배경 색상 변경
        bg_menu = tk.Menu(opt_menu, tearoff=0, font=dropdown_font)
        bg_menu.add_command(label=" 투명 (기본)", image=self.create_menu_color_icon(None, "transparent"), compound="left", command=lambda: self.apply_preset_bg(None))
        bg_menu.add_separator()
        presets = [("검은색", (0,0,0)), ("흰색", (255,255,255)), ("회색", (128,128,128)), ("빨간색", (239,68,68)), ("파란색", (59,130,246)), ("녹색", (34,197,94)), ("노란색", (250, 204, 21)), ("연한 노란색", (254, 240, 138))]
        for name, rgb in presets:
            bg_menu.add_command(label=f" {name}", image=self.create_menu_color_icon(rgb), compound="left", command=lambda r=rgb: self.apply_preset_bg(r))
        bg_menu.add_separator()
        bg_menu.add_command(label=" 직접 지정...", command=self.set_bg_color)
        opt_menu.add_cascade(label=" 배경 색상 변경", menu=bg_menu)
        
        # 아이콘 색상 변경
        ic_menu = tk.Menu(opt_menu, tearoff=0, font=dropdown_font)
        ic_menu.add_command(label=" 원본색 (기본)", image=self.create_menu_color_icon(None, "rainbow"), compound="left", command=lambda: self.apply_preset_icon(None))
        ic_menu.add_separator()
        for name, rgb in presets:
            ic_menu.add_command(label=f" {name}", image=self.create_menu_color_icon(rgb), compound="left", command=lambda r=rgb: self.apply_preset_icon(r))
        ic_menu.add_separator()
        ic_menu.add_command(label=" 직접 지정...", command=self.set_icon_color)
        opt_menu.add_cascade(label=" 아이콘 색상 변경", menu=ic_menu)
        
        opt_menu.add_separator()
        opt_menu.add_command(label=" ✨ 원본 바탕 투명화 변환", command=self.convert_to_transparent)
        
        menubar.add_cascade(label="변환", menu=opt_menu)
        
        # 스타일 메뉴
        style_menu = tk.Menu(menubar, tearoff=0, font=dropdown_font)
        style_menu.add_radiobutton(label=" 기본", variable=self.style_var, value="Normal", command=self.apply_style)
        style_menu.add_radiobutton(label=" 아이폰 기본 스타일", variable=self.style_var, value="iPhone", command=self.apply_style)
        style_menu.add_radiobutton(label=" 아이폰 글래스 스타일", variable=self.style_var, value="Glass", command=self.apply_style)
        menubar.add_cascade(label="스타일", menu=style_menu)

    def setup_ui_styles(self):
        sub_font = ("맑은 고딕", max(9, self.font_size - 2), "bold")
        
        self.btn_reset_bg.config(font=sub_font)
        self.btn_reset_all.config(font=sub_font)
        self.btn_reset_icon.config(font=sub_font)
        self.lbl_preview.config(font=("맑은 고딕", max(10, self.font_size - 1), "bold"))
        self.status_bar.config(font=("맑은 고딕", max(9, self.font_size - 2)))

    def create_checkerboard(self, size, grid_size=16):
        bg = Image.new('RGBA', size, (255, 255, 255, 255))
        draw = ImageDraw.Draw(bg)
        for y in range(0, size[1], grid_size):
            for x in range(0, size[0], grid_size):
                if (x // grid_size + y // grid_size) % 2 == 1:
                    draw.rectangle([x, y, x + grid_size, y + grid_size], fill=(235, 235, 235, 255))
        return bg

    def set_default_checkerboard(self):
        self.tk_image = ImageTk.PhotoImage(self.fixed_bg_image)
        self.lbl_preview.config(image=self.tk_image, compound="center", text="이미지를 불러와주세요")

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("모든 이미지 파일", "*.png *.ico *.jpg *.jpeg *.gif *.svg"),
            ("PNG 파일 (*.png)", "*.png"),
            ("아이콘 파일 (*.ico)", "*.ico"),
            ("JPEG 파일 (*.jpg;*.jpeg)", "*.jpg *.jpeg"),
            ("GIF 파일 (*.gif)", "*.gif"),
            ("SVG 파일 (*.svg)", "*.svg"),
            ("모든 파일 (*.*)", "*.*")
        ])
        if path:
            self.file_path = path
            self.status_var.set(f"파일 불러옴: {os.path.basename(path)}")
            
            if path.lower().endswith('.ico'):
                try:
                    img = Image.open(path)
                    best_size = (0, 0)
                    best_frame = 0
                    
                    for frame in range(getattr(img, "n_frames", 1)):
                        img.seek(frame)
                        if img.size[0] > best_size[0]:
                            best_size = img.size
                            best_frame = frame
                    
                    img.seek(best_frame)
                    # ⭐ 저해상도 이미지를 불러오더라도 고화질(512x512) 캔버스로 강제 스케일 업 보정
                    extracted_img = img.convert("RGBA")
                    if extracted_img.size[0] < 512 or extracted_img.size[1] < 512:
                        extracted_img = extracted_img.resize((512, 512), Image.Resampling.LANCZOS)
                except Exception:
                    extracted_img = Image.open(path).convert("RGBA").resize((512, 512), Image.Resampling.LANCZOS)
            elif path.lower().endswith('.svg'):
                try:
                    from svglib.svglib import svg2rlg
                    from reportlab.graphics import renderPM
                    drawing = svg2rlg(path)
                    png_data = renderPM.drawToString(drawing, fmt="PNG")
                    extracted_img = Image.open(io.BytesIO(png_data)).convert("RGBA").resize((512, 512), Image.Resampling.LANCZOS)
                except ImportError:
                    messagebox.showerror("라이브러리 필요", "SVG 파일을 불러오려면 svglib 및 reportlab 라이브러리가 필요합니다.\n터미널에 'pip install svglib reportlab'을 실행해 주세요.")
                    return
                except Exception as e:
                    messagebox.showerror("오류", f"SVG 파일을 불러오는 중 오류가 발생했습니다.\n{e}")
                    return
            else:
                extracted_img = Image.open(path).convert("RGBA").resize((512, 512), Image.Resampling.LANCZOS)
                
            self.original_img = extracted_img
            self.source_img = extracted_img.copy()
            
            self.bg_color = None
            self.icon_color = None
            self.is_transparent_mode = False
            self.apply_style()

    def convert_to_transparent(self):
        if not self.original_img: return
        img = self.original_img.convert("RGBA")
        pixels, width, height = img.load(), img.size[0], img.size[1]
        bg_target, tolerance = pixels[0, 0], 25
        for y in range(height):
            for x in range(width):
                r, g, b, a = pixels[x, y]
                if all(abs(c1 - c2) <= tolerance for c1, c2 in zip((r, g, b), bg_target[:3])):
                    pixels[x, y] = (255, 255, 255, 0)
        self.source_img, self.is_transparent_mode = img, True
        self.apply_style()

    def reset_to_original(self):
        if self.original_img:
            self.bg_color, self.icon_color, self.is_transparent_mode = None, None, False
            self.source_img = self.original_img.copy()
            self.style_var.set("Normal")
            self.apply_style()

    def apply_preset_bg(self, color):
        self.bg_color = color
        if color and self.style_var.get() == "Normal": self.style_var.set("iPhone")
        self.apply_style()

    def apply_preset_icon(self, color):
        self.icon_color = color
        self.apply_style()

    def reset_bg_color(self): self.apply_preset_bg(None)
    def reset_icon_color(self): self.apply_preset_icon(None)

    def set_bg_color(self):
        color = colorchooser.askcolor(title="배경색 선택")[0]
        if color: self.apply_preset_bg(tuple(map(int, color)))

    def set_icon_color(self):
        color = colorchooser.askcolor(title="아이콘색 선택")[0]
        if color: self.apply_preset_icon(tuple(map(int, color)))

    def apply_style(self):
        if not self.source_img: return
        mode = self.style_var.get()
        icon_source = PILImageOps.colorize(self.source_img.convert("L"), black=self.icon_color, white=(255,255,255)).convert("RGBA") if self.icon_color else self.source_img.copy()
        
        # ⭐ 마스터 원본 캔버스 해상도는 항상 무손실 512x512 고정
        base_size = (512, 512)
        radius = 512 // 4.5
        
        if mode in ["iPhone", "Glass"]:
            bg = Image.new('RGBA', base_size, self.bg_color + (255,) if self.bg_color else (0,0,0,0))
            bg_mask = Image.new('L', base_size, 0)
            ImageDraw.Draw(bg_mask).rounded_rectangle([(0, 0), base_size], radius=radius, fill=255)
            working_img = Image.new('RGBA', base_size, (0,0,0,0))
            working_img.paste(bg, (0, 0), mask=bg_mask)
            
            fit_size = int(512 * 0.95)
            icon_resized = ImageOps.contain(icon_source, (fit_size, fit_size), Image.Resampling.LANCZOS)
            icon_mask = Image.new('L', icon_resized.size, 0)
            ImageDraw.Draw(icon_mask).rounded_rectangle([(0, 0), icon_resized.size], radius=icon_resized.size[0]//4.5, fill=255)
            icon_rounded = Image.new('RGBA', icon_resized.size, (0,0,0,0))
            icon_rounded.paste(icon_resized, (0, 0), mask=icon_mask)
            
            working_img.paste(icon_rounded, ((512-icon_rounded.size[0])//2, (512-icon_rounded.size[1])//2), mask=icon_rounded)
            
            if mode == "Glass":
                gloss = Image.new('RGBA', base_size, (255, 255, 255, 0))
                g_draw = ImageDraw.Draw(gloss)
                for i in range(256): g_draw.line([(0, i), (512, i)], fill=(255, 255, 255, int(85 * (1 - i/256))))
                gloss_masked = Image.new('RGBA', base_size, (0,0,0,0))
                gloss_masked.paste(gloss, (0, 0), mask=bg_mask)
                working_img = Image.alpha_composite(working_img, gloss_masked)
                
                ol_layer = Image.new('RGBA', base_size, (0,0,0,0))
                edge = (0,0,0,55) if self.bg_color and sum(self.bg_color)/3 > 180 else (255,255,255,80)
                ImageDraw.Draw(ol_layer).rounded_rectangle([(0, 0), base_size], radius=radius, outline=edge, width=3)
                working_img = Image.alpha_composite(working_img, ol_layer)
        else:
            working_img = ImageOps.contain(icon_source, base_size, Image.Resampling.LANCZOS)

        self.current_img = working_img
        self.update_preview()

    def update_preview(self):
        if not self.current_img: return
        canvas = self.fixed_bg_image.copy()
        
        # ⭐ 미리보기를 그릴 때 고해상도(512x512) 마스터 원본에서 즉시 선명하게 다운스케일 렌더링
        display_img = ImageOps.contain(self.current_img, (280, 280), Image.Resampling.LANCZOS)
        temp_layer = Image.new('RGBA', (296, 296), (0,0,0,0))
        temp_layer.paste(display_img, ((296-display_img.size[0])//2, (296-display_img.size[1])//2))
        
        self.tk_image = ImageTk.PhotoImage(Image.alpha_composite(canvas, temp_layer))
        self.lbl_preview.config(image=self.tk_image, text="")

    def save_file(self):
        if not self.current_img: return
        
        today_str = datetime.now().strftime("%y%m%d")
        current_style = self.style_var.get()
        
        if self.file_path:
            base_name = os.path.basename(self.file_path)
            file_name_without_ext, _ = os.path.splitext(base_name)
            suggested_filename = f"{file_name_without_ext}_{current_style}_{today_str}.ico"
        else:
            suggested_filename = f"icon_{current_style}_{today_str}.ico"

        path = filedialog.asksaveasfilename(
            defaultextension=".ico", 
            initialfile=suggested_filename,
            filetypes=[("Icon files", "*.ico")]
        )
        
        if path:
            max_s = self.icon_size.get()
            # ⭐ 중요: 내보내기용 이미지 압축 어레이를 생성할 때 복사본(.copy())을 취해 원본 인스턴스 보호
            save_canvas = self.current_img.copy()
            
            # ICO 표준 멀티 레이어 빌드업
            layers = [save_canvas.resize((s, s), Image.Resampling.LANCZOS) for s in [16, 32, 48, 64, 128, 256, 512] if s <= max_s]
            
            # 해상도가 높은 순서대로 빌드되도록 역순 정렬 처리
            layers.sort(key=lambda x: x.size[0], reverse=True)
            
            layers[0].save(
                path, 
                format="ICO", 
                sizes=[(l.size[0], l.size[1]) for l in layers], 
                append_images=layers[1:]
            )
            
            # ⭐ 파일 저장 후에도 원본 컨텍스트 파괴 없이 즉시 미리보기 재호출하여 선명한 화면 강제 유지
            self.update_preview()
            messagebox.showinfo("완료", f"아이콘이 저장되었습니다.\n파일명: {os.path.basename(path)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = IconGeneratorApp(root)
    root.mainloop()