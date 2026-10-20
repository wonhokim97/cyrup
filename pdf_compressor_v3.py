"""PDF 용량 줄이기: PyMuPDF + Tkinter. 실행: python pdf_compressor.py"""
import os
import queue
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import pymupdf as fitz

class CompressionCancelled(Exception):
    """압축 작업 취소."""

PRESETS = {
    "무손실 (원본 화질 유지)": None,
    "균형 (이미지 최적화)": (180, 75),
    "강력 (작은 파일 우선)": (120, 55),
}


def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{n} B"
        n /= 1024


def compress_pdf(source, destination, preset, password="", cancel_event=None, progress=None):
    """임시 파일 작성 후 성공했을 때만 목적지 교체. 실패 시 원본/목적지 유지."""
    if os.path.normcase(os.path.abspath(source)) == os.path.normcase(os.path.abspath(destination)):
        raise ValueError("원본과 저장 경로는 달라야 합니다.")
    def checkpoint(stage):
        if cancel_event is not None and cancel_event.is_set():
            raise CompressionCancelled("작업이 취소되었습니다.")
        if progress is not None:
            progress(stage)

    temp_path = None
    doc = None
    try:
        checkpoint("PDF 열기")
        doc = fitz.open(source)
        if doc.needs_pass and not doc.authenticate(password):
            raise ValueError("PDF 비밀번호가 올바르지 않습니다.")
        mode = PRESETS[preset]
        if mode:
            dpi, quality = mode
            if not hasattr(doc, "rewrite_images"):
                raise RuntimeError("이미지 최적화에는 PyMuPDF 1.26 이상이 필요합니다.")
            checkpoint("이미지 최적화")
            doc.rewrite_images(
                dpi_threshold=dpi + 1,
                dpi_target=dpi,
                quality=quality,
                lossy=True, lossless=True,
                bitonal=False, color=True, gray=True,
            )
        checkpoint("PDF 저장")
        folder = os.path.dirname(os.path.abspath(destination))
        fd, temp_path = tempfile.mkstemp(prefix=".pdf_compress_", suffix=".pdf", dir=folder)
        os.close(fd)
        doc.save(
            temp_path,
            garbage=4,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            use_objstms=True,
            encryption=fitz.PDF_ENCRYPT_KEEP,  # pyright: ignore[reportAttributeAccessIssue]
        )

        doc.close()
        doc = None
        checkpoint("결과 검증")
        # 출력을 다시 열어 유효성 확인 후 기존 목적지를 교체합니다.
        with fitz.open(temp_path) as check:
            if check.needs_pass:
                if not check.authenticate(password):
                    raise RuntimeError("저장된 PDF를 검증할 수 없습니다.")
            if check.page_count < 1:
                raise RuntimeError("저장된 PDF에 페이지가 없습니다.")
        size = os.path.getsize(temp_path)
        checkpoint("저장 완료 처리")
        os.replace(temp_path, destination)
        temp_path = None
        return size
    finally:
        if doc is not None:
            doc.close()
        if temp_path is not None and os.path.exists(temp_path):
            os.unlink(temp_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 용량 줄이기")
        self.geometry("920x740")
        self.minsize(640, 480)
        self.doc = None
        self.source = ""
        self.password = ""
        self.page_index = 0
        self.photo = None
        self.events = queue.Queue()
        self.busy = False
        self.cancel_event = None
        self.notify_complete = tk.BooleanVar(value=True)
        self._build()
        self.bind_all("<Control-o>", lambda event: self.choose())
        self.bind_all("<Control-s>", lambda event: self.start())
        self.bind_all("<Control-Left>", lambda event: self.turn(-1))
        self.bind_all("<Control-Right>", lambda event: self.turn(1))
        self.bind_all("<Escape>", lambda event: self.cancel() if self.busy else None)
        self.after(120, self._poll)
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _log(self, text):
        self.system_log.config(text="System Log: " + text)

    def _build(self):
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="PDF 파일 선택", accelerator="Ctrl+O", command=self.choose)
        file_menu.add_command(label="압축하여 저장", accelerator="Ctrl+S", command=self.start)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self._close)
        menu.add_cascade(label="파일", menu=file_menu)
        view_menu = tk.Menu(menu, tearoff=False)
        view_menu.add_command(label="이전 페이지", accelerator="Ctrl+←", command=lambda: self.turn(-1))
        view_menu.add_command(label="다음 페이지", accelerator="Ctrl+→", command=lambda: self.turn(1))
        menu.add_cascade(label="보기", menu=view_menu)
        compress_menu = tk.Menu(menu, tearoff=False)
        compress_menu.add_command(label="압축하여 저장", accelerator="Ctrl+S", command=self.start)
        compress_menu.add_command(label="작업 취소", accelerator="Esc", command=self.cancel)
        compress_menu.add_separator()
        compress_menu.add_checkbutton(label="압축 완료 알림 표시", variable=self.notify_complete)
        menu.add_cascade(label="압축", menu=compress_menu)
        self.config(menu=menu)

        tools = ttk.Frame(self, padding=(12, 8))
        tools.pack(fill="x")
        ttk.Button(tools, text="📂 파일 선택", command=self.choose).pack(side="left")
        self.toolbar_save = ttk.Button(tools, text="💾 압축 저장", command=self.start)
        self.toolbar_save.pack(side="left", padx=5)
        self.toolbar_cancel = ttk.Button(tools, text="✖ 작업 취소", command=self.cancel, state="disabled")
        self.toolbar_cancel.pack(side="left")
        ttk.Button(tools, text="◀ 이전", command=lambda: self.turn(-1)).pack(side="left", padx=(18, 0))
        ttk.Button(tools, text="다음 ▶", command=lambda: self.turn(1)).pack(side="left", padx=5)

        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")
        ttk.Button(top, text="PDF 파일 선택", command=self.choose).pack(side="left")
        self.filename = ttk.Label(top, text="선택된 파일 없음", width=60)
        self.filename.pack(side="left", padx=12, fill="x", expand=True)

        bar = ttk.Frame(self, padding=(12, 0, 12, 8))
        bar.pack(fill="x")
        ttk.Label(bar, text="압축 수준").pack(side="left")
        self.preset = tk.StringVar(value=next(iter(PRESETS)))
        ttk.Combobox(bar, textvariable=self.preset, values=list(PRESETS),
                     state="readonly", width=26).pack(side="left", padx=8)
        self.compress_button = ttk.Button(bar, text="압축하여 저장", command=self.start)
        self.compress_button.pack(side="right")
        ttk.Label(bar, text="이미지 최적화는 화질이 낮아질 수 있습니다.").pack(side="left", padx=10)

        self.info = ttk.Label(self, text="PDF를 선택하면 첫 페이지가 표시됩니다.", padding=(12, 0))
        self.info.pack(fill="x")
        preview_area = ttk.Frame(self, padding=12)
        preview_area.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(preview_area, bg="#555555", highlightthickness=0)
        vs = ttk.Scrollbar(preview_area, orient="vertical", command=self.canvas.yview)
        hs = ttk.Scrollbar(preview_area, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        preview_area.rowconfigure(0, weight=1)
        preview_area.columnconfigure(0, weight=1)
        self.canvas.bind("<Configure>", self._resize_preview)

        foot = ttk.Frame(self, padding=(12, 0, 12, 5))
        foot.pack(fill="x")
        ttk.Button(foot, text="◀ 이전", command=lambda: self.turn(-1)).pack(side="left")
        ttk.Button(foot, text="다음 ▶", command=lambda: self.turn(1)).pack(side="left", padx=6)
        self.page_label = ttk.Label(foot, text="페이지 - / -")
        self.page_label.pack(side="left", padx=12)
        self.status = ttk.Label(foot, text="대기 중")
        self.status.pack(side="right")
        progress_row = ttk.Frame(self, padding=(12, 0, 12, 4))
        progress_row.pack(fill="x")
        self.progress = ttk.Progressbar(progress_row, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True)
        self.stage = ttk.Label(progress_row, text="대기 중", width=18, anchor="e")
        self.stage.pack(side="right", padx=(8, 0))
        self.system_log = ttk.Label(self, text="System Log: 대기 중", padding=(12, 2), anchor="w")
        self.system_log.pack(side="bottom", fill="x")

    def choose(self):
        if self.busy:
            return
        path = filedialog.askopenfilename(title="PDF 파일 선택", filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")])
        if not path:
            return
        candidate = None
        try:
            candidate = fitz.open(path)
            password = ""
            if candidate.needs_pass:
                password = simpledialog.askstring("비밀번호", "PDF 비밀번호를 입력하세요:", show="*", parent=self)
                if password is None:
                    candidate.close()
                    return
                if not candidate.authenticate(password):
                    raise ValueError("PDF 비밀번호가 올바르지 않습니다.")
            if candidate.page_count < 1:
                raise ValueError("페이지가 없는 PDF입니다.")
            if self.doc:
                self.doc.close()
            self.doc, candidate = candidate, None
            self.source, self.password, self.page_index = path, password, 0
            self.filename.config(text=os.path.basename(path))
            self.info.config(text=f"원본 크기: {human_size(os.path.getsize(path))}  |  총 {self.doc.page_count}페이지")
            self.render()
            self._log("파일 선택: " + os.path.basename(path))
        except Exception as exc:
            messagebox.showerror("파일 열기 실패", str(exc), parent=self)
        finally:
            if candidate:
                candidate.close()

    def _resize_preview(self, _event):
        if self.doc:
            if hasattr(self, "_resize_job"):
                self.after_cancel(self._resize_job)
            self._resize_job = self.after(180, self.render)

    def render(self):
        if not self.doc:
            return
        try:
            page = self.doc.load_page(self.page_index)
            rect = page.rect
            width = max(200, self.canvas.winfo_width() - 28)
            height = max(200, self.canvas.winfo_height() - 28)
            scale = min(width / rect.width, height / rect.height, 2)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            self.photo = tk.PhotoImage(data=pix.tobytes("ppm"))
            self.canvas.delete("all")
            x = max(14, (self.canvas.winfo_width() - pix.width) // 2)
            self.canvas.create_image(x, 14, anchor="nw", image=self.photo)
            self.canvas.configure(scrollregion=(0, 0, max(self.canvas.winfo_width(), x + pix.width + 14), pix.height + 28))
            self.page_label.config(text=f"페이지 {self.page_index + 1} / {self.doc.page_count}")
        except Exception as exc:
            self.status.config(text=f"미리보기 오류: {exc}")

    def turn(self, delta):
        if self.doc:
            index = max(0, min(self.doc.page_count - 1, self.page_index + delta))
            if index != self.page_index:
                self.page_index = index
                self.render()

    def start(self):
        if self.busy:
            return
        if not self.source:
            messagebox.showwarning("파일 선택", "먼저 PDF 파일을 선택하세요.", parent=self)
            return
        name, _ = os.path.splitext(self.source)
        dest = filedialog.asksaveasfilename(
            title="압축 PDF 저장", defaultextension=".pdf", initialfile=os.path.basename(name) + "_compressed.pdf",
            initialdir=os.path.dirname(self.source), filetypes=[("PDF 파일", "*.pdf")])
        if not dest:
            return
        if os.path.normcase(os.path.abspath(dest)) == os.path.normcase(os.path.abspath(self.source)):
            messagebox.showerror("저장 경로 오류", "원본 파일과 다른 이름으로 저장하세요.", parent=self)
            return
        self.busy = True
        self.cancel_event = threading.Event()
        self.compress_button.config(state="disabled")
        self.toolbar_save.config(state="disabled")
        self.toolbar_cancel.config(state="normal")
        self.status.config(text="압축 중...")
        self.stage.config(text="작업 시작")
        self._log("압축 시작")
        self.progress.start(12)
        source, password, preset = self.source, self.password, self.preset.get()
        threading.Thread(target=self._worker, args=(source, dest, preset, password, self.cancel_event), daemon=True).start()

    def cancel(self):
        cancel_event = self.cancel_event
        if self.busy and cancel_event is not None and not cancel_event.is_set():
            cancel_event.set()
            self.stage.config(text="취소 대기 중")
            self.status.config(text="취소 요청 중...")
            self._log("취소 요청: 현재 처리 단계가 끝나면 중단합니다.")

    def _worker(self, source, dest, preset, password, cancel_event):
        try:
            before = os.path.getsize(source)
            after = compress_pdf(
                source, dest, preset, password, cancel_event,
                lambda stage: self.events.put(("progress", stage)),
            )
            self.events.put(("ok", before, after, dest))
        except CompressionCancelled:
            self.events.put(("cancelled",))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    cancel_event = self.cancel_event
                    if self.busy and cancel_event is not None and not cancel_event.is_set():
                        self.stage.config(text=event[1])
                        self._log(event[1] + " 중...")
                    continue

                self.busy = False
                self.progress.stop()
                self.compress_button.config(state="normal")
                self.toolbar_save.config(state="normal")
                self.toolbar_cancel.config(state="disabled")
                if event[0] == "ok":
                    _, before, after, dest = event
                    difference = before - after
                    change = difference / before * 100 if before else 0
                    text = (
                        f"저장 완료: {dest}\n"
                        f"원본 {human_size(before)} → 결과 {human_size(after)}\n"
                        f"{'감소' if difference >= 0 else '증가'} "
                        f"{human_size(abs(difference))} ({abs(change):.1f}%)"
                    )

                    self.stage.config(text="완료")
                    self.status.config(text="저장 완료")
                    self._log(text.replace("\n", " | "))
                    if self.notify_complete.get():
                        messagebox.showinfo("압축 결과", text, parent=self)
                elif event[0] == "cancelled":
                    self.stage.config(text="취소됨")
                    self.status.config(text="작업 취소")
                    self._log("압축이 취소되었습니다. 기존 저장 파일은 변경되지 않았습니다.")
                else:
                    self.stage.config(text="실패")
                    self.status.config(text="압축 실패")
                    self._log("압축 실패: " + event[1])
                    messagebox.showerror("압축 실패", event[1], parent=self)
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _close(self):
        if self.busy:
            messagebox.showinfo("압축 중", "압축 작업이 끝난 후 창을 닫아 주세요.", parent=self)
            return
        if self.doc:
            self.doc.close()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
