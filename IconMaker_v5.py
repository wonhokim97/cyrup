import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageOps, ImageOps as PILImageOps
import os
import io

class IconGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional Icon Generator")
        
        # 기본 기준 사이즈 설정
        self.base_width = 490
        self.base_height = 460
        self.root.geometry(f"{self.base_width}x{self.base_height}")
        self.root.resizable(True, True)

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
        
        self.current_font_size = tk.IntVar(value=11)
        self.active_dropdown = None
        self.active_menu_button = None
        
        # ⭐ 메뉴 아이콘 이미지 가비지 컬렉션 방지용 리스트
        self.menu_icons_refs = []

        # 메뉴바 프레임
        self.menubar_frame = tk.Frame(self.root, bg="#f0f0f0", bd=1, relief=tk.RAISED)
        self.menubar_frame.pack(side=tk.TOP, fill=tk.X)
        
        self.menu_buttons = []
        self.setup_custom_menubar()
        
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(expand=True, fill=tk.BOTH, pady=10)
        
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

        self.root.bind("<Button-1>", self.close_dropdowns)
        self.update_ui_fonts_and_scale()

    # ⭐ 메뉴에 들어갈 실제 컬러 사각형 아이콘 생성 함수
    def create_menu_color_icon(self, color_tuple, special=None):
        size = int(self.current_font_size.get() * 1.2)
        if special == "transparent":
            # 투명 아이콘: 격자무늬
            img = self.create_checkerboard((size, size), grid_size=size//2)
        elif special == "rainbow":
            # 원본색 아이콘: 무지개색
            img = Image.new('RGBA', (size, size), (0,0,0,0))
            draw = ImageDraw.Draw(img)
            draw.rectangle([0, 0, size//2, size//2], fill=(255,0,0,255))
            draw.rectangle([size//2, 0, size, size//2], fill=(0,255,0,255))
            draw.rectangle([0, size//2, size//2, size], fill=(0,0,255,255))
            draw.rectangle([size//2, size//2, size, size], fill=(255,255,0,255))
        else:
            # 일반 단색 아이콘
            img = Image.new('RGBA', (size, size), color_tuple + (255,))
            ImageDraw.Draw(img).rectangle([0, 0, size-1, size-1], outline=(150,150,150,255)) # 미세 테두리

        tk_img = ImageTk.PhotoImage(img)
        self.menu_icons_refs.append(tk_img) # 참조 유지
        return tk_img

    def setup_custom_menubar(self):
        for btn in self.menu_buttons: btn.destroy()
        self.menu_buttons.clear()

        for name, m_type in [("파일", "file"), ("변경", "opt"), ("스타일", "style"), ("환경", "ui_set")]:
            btn = tk.Button(self.menubar_frame, text=name, bg="#f0f0f0", bd=0, padx=15, pady=2, activebackground="#3b82f6", activeforeground="#ffffff")
            btn.pack(side=tk.LEFT)
            btn.bind("<Button-1>", lambda e, t=m_type, b=btn: self.show_dropdown(e, t, b))
            self.menu_buttons.append(btn)

    def update_ui_fonts_and_scale(self):
        size = self.current_font_size.get()
        scale_ratio = size / 11.0
        self.root.geometry(f"{int(self.base_width * scale_ratio)}x{int(self.base_height * scale_ratio)}")
        
        main_font = ("맑은 고딕", size, "bold")
        sub_font = ("맑은 고딕", max(9, size - 2), "bold")
        
        for btn in self.menu_buttons:
            btn.config(font=main_font)
            
        self.btn_reset_bg.config(font=sub_font)
        self.btn_reset_all.config(font=sub_font)
        self.btn_reset_icon.config(font=sub_font)
        self.lbl_preview.config(font=("맑은 고딕", max(10, size - 1), "bold"))
        self.status_bar.config(font=("맑은 고딕", max(9, size - 2)))

    def show_dropdown(self, event, menu_type, button_widget):
        self.close_dropdowns()
        self.menu_icons_refs = [] # 아이콘 리스트 초기화
        
        self.active_menu_button = button_widget
        self.active_menu_button.config(bg="#3b82f6", fg="#ffffff")
        
        size = self.current_font_size.get()
        dropdown_font = ("맑은 고딕", size)
        dropdown = tk.Menu(self.root, tearoff=0, font=dropdown_font)
        
        if menu_type == "file":
            dropdown.add_command(label=" 파일 불러오기", command=self.load_file)
            dropdown.add_command(label=" 다른 이름으로 저장", command=self.save_file)
            dropdown.add_separator()
            dropdown.add_command(label=" 프로그램 종료", command=self.root.quit)
            
        elif menu_type == "opt":
            res_menu = tk.Menu(dropdown, tearoff=0, font=dropdown_font)
            for s in [64, 128, 256, 512]:
                res_menu.add_radiobutton(label=f"{s}x{s} px", variable=self.icon_size, value=s)
            dropdown.add_cascade(label=" 아이콘 크기 선택", menu=res_menu)
            
            # ⭐ 배경 색상 메뉴 (실제 컬러 아이콘 적용)
            bg_menu = tk.Menu(dropdown, tearoff=0, font=dropdown_font)
            bg_menu.add_command(label=" 투명 (기본)", image=self.create_menu_color_icon(None, "transparent"), compound="left", command=lambda: self.apply_preset_bg(None))
            bg_menu.add_separator()
            presets = [("검은색", (0,0,0)), ("흰색", (255,255,255)), ("회색", (128,128,128)), ("빨간색", (239,68,68)), ("파란색", (59,130,246)), ("녹색", (34,197,94))]
            for name, rgb in presets:
                bg_menu.add_command(label=f" {name}", image=self.create_menu_color_icon(rgb), compound="left", command=lambda r=rgb: self.apply_preset_bg(r))
            bg_menu.add_separator()
            bg_menu.add_command(label=" 직접 지정...", command=self.set_bg_color)
            dropdown.add_cascade(label=" 배경 색상 변경", menu=bg_menu)
            
            # ⭐ 아이콘 색상 메뉴 (실제 컬러 아이콘 적용)
            ic_menu = tk.Menu(dropdown, tearoff=0, font=dropdown_font)
            ic_menu.add_command(label=" 원본색 (기본)", image=self.create_menu_color_icon(None, "rainbow"), compound="left", command=lambda: self.apply_preset_icon(None))
            ic_menu.add_separator()
            for name, rgb in presets:
                ic_menu.add_command(label=f" {name}", image=self.create_menu_color_icon(rgb), compound="left", command=lambda r=rgb: self.apply_preset_icon(r))
            ic_menu.add_separator()
            ic_menu.add_command(label=" 직접 지정...", command=self.set_icon_color)
            dropdown.add_cascade(label=" 아이콘 색상 변경", menu=ic_menu)
            
            dropdown.add_separator()
            dropdown.add_command(label=" ✨ 원본 바탕 투명화 변환", command=self.convert_to_transparent)
            
        elif menu_type == "style":
            dropdown.add_radiobutton(label=" 기본", variable=self.style_var, value="Normal", command=self.apply_style)
            dropdown.add_radiobutton(label=" 아이폰 기본 스타일", variable=self.style_var, value="iPhone", command=self.apply_style)
            dropdown.add_radiobutton(label=" 아이폰 글래스 스타일", variable=self.style_var, value="Glass", command=self.apply_style)

        elif menu_type == "ui_set":
            for f_size in range(10, 19):
                dropdown.add_radiobutton(label=f"{f_size} pt", variable=self.current_font_size, value=f_size, command=self.update_ui_fonts_and_scale)

        x, y = event.widget.winfo_rootx(), event.widget.winfo_rooty() + event.widget.winfo_height()
        dropdown.post(x, y)
        self.active_dropdown = dropdown
        return "break"

    def close_dropdowns(self, event=None):
        if self.active_menu_button:
            self.active_menu_button.config(bg="#f0f0f0", fg="#000000") 
            self.active_menu_button = None
        if self.active_dropdown:
            self.active_dropdown.unpost()
            self.active_dropdown = None

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
        path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.ico *.jpg *.jpeg *.gif")])
        if path:
            self.file_path = path
            img = Image.open(path).convert("RGBA")
            self.original_img = img
            self.source_img = img.copy()
            self.apply_style()

    def convert_to_transparent(self):
        if self.original_img is None:
            return

        img = self.original_img.convert("RGBA")
        data = bytearray(img.tobytes())
        tolerance = 25

        bg_r, bg_g, bg_b = data[0], data[1], data[2]

        for i in range(0, len(data), 4):
            if (
                abs(data[i] - bg_r) <= tolerance
                and abs(data[i + 1] - bg_g) <= tolerance
                and abs(data[i + 2] - bg_b) <= tolerance
            ):
                data[i:i + 4] = bytes((255, 255, 255, 0))

        self.source_img = Image.frombytes(
            "RGBA", img.size, bytes(data)
        )
        self.is_transparent_mode = True
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
        current_img = self.current_img
        if current_img is None:
            self.set_default_checkerboard()
            return

        canvas = self.fixed_bg_image.copy()
        display_img = ImageOps.contain(
            current_img,
            (280, 280),
            Image.Resampling.LANCZOS
        )

        temp_layer = Image.new(
            "RGBA", (296, 296), (0, 0, 0, 0)
        )
        position = (
            (296 - display_img.width) // 2,
            (296 - display_img.height) // 2
        )
        temp_layer.paste(display_img, position)

        self.tk_image = ImageTk.PhotoImage(
            Image.alpha_composite(canvas, temp_layer)
        )
        self.lbl_preview.config(image=self.tk_image, text="")

    def save_file(self):
        if not self.current_img: return
        path = filedialog.asksaveasfilename(defaultextension=".ico", filetypes=[("Icon files", "*.ico")])
        if path:
            max_s = self.icon_size.get()
            layers = [self.current_img.resize((s, s), Image.Resampling.LANCZOS) for s in [16, 32, 48, 64, 128, 256, 512] if s <= max_s]
            layers[0].save(path, format="ICO", sizes=[(l.size[0], l.size[1]) for l in layers], append_images=layers[1:])
            messagebox.showinfo("완료", "아이콘이 저장되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = IconGeneratorApp(root)
    root.mainloop()