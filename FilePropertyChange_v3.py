import os
import sys
import shutil
import ctypes
import tempfile
import tkinter as tk

from datetime import datetime
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter


SUPPORTED_EXTENSIONS = (".pdf", ".txt", ".csv")
DATE_FORMAT = "%Y%m%d%H%M%S"

row_widgets = {}


# ==================================================
# 날짜 변환
# ==================================================
def format_timestamp(timestamp):
    try:
        return datetime.fromtimestamp(timestamp).strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    except (ValueError, OSError, TypeError):
        return "정보 없음"


def timestamp_to_input(timestamp):
    try:
        return datetime.fromtimestamp(timestamp).strftime(DATE_FORMAT)
    except (ValueError, OSError, TypeError):
        return ""


def format_pdf_date(raw_date):
    if not raw_date:
        return "정보 없음"

    clean_date = str(raw_date)

    if clean_date.startswith("D:"):
        clean_date = clean_date[2:]

    if len(clean_date) >= 14 and clean_date[:14].isdigit():
        return (
            f"{clean_date[0:4]}-{clean_date[4:6]}-{clean_date[6:8]} "
            f"{clean_date[8:10]}:{clean_date[10:12]}:{clean_date[12:14]}"
        )

    return str(raw_date)


def extract_pdf_date(raw_date):
    if not raw_date:
        return ""

    clean_date = str(raw_date)

    if clean_date.startswith("D:"):
        clean_date = clean_date[2:]

    if len(clean_date) >= 14 and clean_date[:14].isdigit():
        return clean_date[:14]

    return ""


def parse_input_date(date_text):
    if len(date_text) != 14 or not date_text.isdigit():
        raise ValueError(
            "날짜를 YYYYMMDDHHMMSS 형식의 "
            "14자리 숫자로 입력해 주세요."
        )

    try:
        return datetime.strptime(date_text, DATE_FORMAT)
    except ValueError:
        raise ValueError(
            f"실제로 존재하는 날짜를 입력해 주세요.\n"
            f"입력값: {date_text}"
        )


def get_creation_timestamp(file_stat):
    if sys.platform == "win32":
        return file_stat.st_ctime

    if hasattr(file_stat, "st_birthtime"):
        return file_stat.st_birthtime

    return file_stat.st_ctime


# ==================================================
# Windows 파일 날짜 변경
# ==================================================
def set_windows_file_times(
    file_path,
    creation_datetime,
    access_datetime,
    modified_datetime
):
    if sys.platform != "win32":
        raise OSError("Windows에서만 날짜를 변경할 수 있습니다.")

    FILE_WRITE_ATTRIBUTES = 0x0100
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x00000080
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

    class FILETIME(ctypes.Structure):
        _fields_ = [
            ("dwLowDateTime", ctypes.c_uint32),
            ("dwHighDateTime", ctypes.c_uint32),
        ]

    def convert_to_filetime(date_value):
        unix_timestamp = date_value.timestamp()

        windows_timestamp = int(
            (unix_timestamp + 11644473600) * 10_000_000
        )

        return FILETIME(
            windows_timestamp & 0xFFFFFFFF,
            windows_timestamp >> 32
        )

    creation_time = convert_to_filetime(creation_datetime)
    access_time = convert_to_filetime(access_datetime)
    modified_time = convert_to_filetime(modified_datetime)

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        ctypes.c_wchar_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
    ]
    create_file.restype = ctypes.c_void_p

    set_file_time = kernel32.SetFileTime
    set_file_time.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
    ]
    set_file_time.restype = ctypes.c_int

    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [ctypes.c_void_p]
    close_handle.restype = ctypes.c_int

    handle = create_file(
        os.path.abspath(file_path),
        FILE_WRITE_ATTRIBUTES,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None
    )

    if handle == INVALID_HANDLE_VALUE:
        raise ctypes.WinError(ctypes.get_last_error())

    try:
        result = set_file_time(
            handle,
            ctypes.byref(creation_time),
            ctypes.byref(access_time),
            ctypes.byref(modified_time)
        )

        if result == 0:
            raise ctypes.WinError(ctypes.get_last_error())

    finally:
        close_handle(handle)


# ==================================================
# 날짜 정보 행 관리
# ==================================================
def set_row_value(key, current_text, input_text, enabled=True):
    row = row_widgets[key]

    row["current"].config(
        text=current_text,
        fg="#1976D2" if enabled else "gray"
    )

    row["value"].set(input_text)
    row["selected"].set(False)

    state = "normal" if enabled else "disabled"

    row["check"].config(state=state)
    row["entry"].config(state=state)


def clear_information():
    lbl_file_type.config(text="-", fg="gray")

    for row in row_widgets.values():
        row["current"].config(text="-", fg="gray")
        row["value"].set("")
        row["selected"].set(False)
        row["check"].config(state="disabled")
        row["entry"].config(state="disabled")


def select_all_available():
    for row in row_widgets.values():
        if str(row["check"].cget("state")) == "normal":
            row["selected"].set(True)


def clear_all_selections():
    for row in row_widgets.values():
        row["selected"].set(False)


# ==================================================
# 파일 선택
# ==================================================
def select_file():
    file_path = filedialog.askopenfilename(
        title="변경할 파일 선택",
        filetypes=[
            ("지원 파일", "*.pdf *.txt *.csv"),
            ("PDF 파일", "*.pdf"),
            ("텍스트 파일", "*.txt"),
            ("CSV 파일", "*.csv"),
            ("모든 파일", "*.*"),
        ]
    )

    if not file_path:
        return

    entry_file.delete(0, tk.END)
    entry_file.insert(0, file_path)

    load_file_information(file_path)


# ==================================================
# 파일 정보 불러오기
# ==================================================
def load_file_information(file_path):
    clear_information()

    if not os.path.isfile(file_path):
        messagebox.showerror(
            "오류",
            "선택한 파일을 찾을 수 없습니다."
        )
        return

    extension = os.path.splitext(file_path)[1].lower()

    if extension not in SUPPORTED_EXTENSIONS:
        messagebox.showerror(
            "오류",
            "PDF, TXT, CSV 파일만 선택할 수 있습니다."
        )
        return

    saved_stat = None

    try:
        saved_stat = os.stat(file_path)

        creation_timestamp = get_creation_timestamp(saved_stat)
        access_timestamp = saved_stat.st_atime
        modified_timestamp = saved_stat.st_mtime

        lbl_file_type.config(
            text=extension[1:].upper(),
            fg="#1565C0"
        )

        set_row_value(
            "fs_creation",
            format_timestamp(creation_timestamp),
            timestamp_to_input(creation_timestamp)
        )

        set_row_value(
            "fs_access",
            format_timestamp(access_timestamp),
            timestamp_to_input(access_timestamp)
        )

        set_row_value(
            "fs_modified",
            format_timestamp(modified_timestamp),
            timestamp_to_input(modified_timestamp)
        )

        if extension == ".pdf":
            load_pdf_information(file_path)
        else:
            set_row_value(
                "pdf_creation",
                "해당 없음",
                "",
                enabled=False
            )

            set_row_value(
                "pdf_modified",
                "해당 없음",
                "",
                enabled=False
            )

    except Exception as error:
        clear_information()

        messagebox.showerror(
            "오류",
            f"파일 정보를 읽을 수 없습니다.\n\n{error}"
        )

    finally:
        # 파일 조회로 액세스 날짜가 변경되는 것을 방지
        if (
            saved_stat is not None
            and os.path.isfile(file_path)
            and sys.platform == "win32"
        ):
            try:
                set_windows_file_times(
                    file_path,
                    datetime.fromtimestamp(
                        get_creation_timestamp(saved_stat)
                    ),
                    datetime.fromtimestamp(saved_stat.st_atime),
                    datetime.fromtimestamp(saved_stat.st_mtime)
                )
            except Exception:
                pass


def load_pdf_information(file_path):
    with open(file_path, "rb") as pdf_file:
        reader = PdfReader(pdf_file)
        metadata = reader.metadata or {}

        creation_raw = metadata.get("/CreationDate")
        modified_raw = metadata.get("/ModDate")

    set_row_value(
        "pdf_creation",
        format_pdf_date(creation_raw),
        extract_pdf_date(creation_raw)
    )

    set_row_value(
        "pdf_modified",
        format_pdf_date(modified_raw),
        extract_pdf_date(modified_raw)
    )


# ==================================================
# PDF 원본 수정
# ==================================================
def modify_pdf_in_place(
    file_path,
    pdf_creation_text,
    pdf_modified_text,
    original_creation,
    original_access,
    original_modified,
    target_creation,
    target_access,
    target_modified
):
    directory = os.path.dirname(os.path.abspath(file_path))

    temp_path = None
    backup_path = None
    completed = False

    try:
        with open(file_path, "rb") as source_file:
            reader = PdfReader(source_file)
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            metadata = {}

            for key, value in dict(reader.metadata or {}).items():
                if value is not None:
                    metadata[str(key)] = str(value)

            if pdf_creation_text is not None:
                metadata["/CreationDate"] = (
                    f"D:{pdf_creation_text}+09'00'"
                )

            if pdf_modified_text is not None:
                metadata["/ModDate"] = (
                    f"D:{pdf_modified_text}+09'00'"
                )

            writer.add_metadata(metadata)

            with tempfile.NamedTemporaryFile(
                mode="wb",
                suffix=".pdf",
                prefix="pdf_temp_",
                dir=directory,
                delete=False
            ) as temp_file:
                temp_path = temp_file.name
                writer.write(temp_file)

        try:
            shutil.copymode(file_path, temp_path)
        except OSError:
            pass

        # 원본 백업 생성
        backup_fd, backup_path = tempfile.mkstemp(
            suffix=".bak",
            prefix="pdf_backup_",
            dir=directory
        )
        os.close(backup_fd)

        shutil.copy2(file_path, backup_path)

        # 임시 파일로 원본 교체
        os.replace(temp_path, file_path)
        temp_path = None

        set_windows_file_times(
            file_path,
            target_creation,
            target_access,
            target_modified
        )

        completed = True

    except Exception as original_error:
        restore_error = None

        if backup_path and os.path.exists(backup_path):
            try:
                os.replace(backup_path, file_path)
                backup_path = None

                set_windows_file_times(
                    file_path,
                    original_creation,
                    original_access,
                    original_modified
                )

            except Exception as error:
                restore_error = error

        if restore_error:
            raise RuntimeError(
                "PDF 변경 및 원본 복원에 실패했습니다.\n\n"
                f"변경 오류: {original_error}\n"
                f"복원 오류: {restore_error}\n"
                f"백업 위치: {backup_path}"
            )

        raise original_error

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass

        if completed and backup_path and os.path.exists(backup_path):
            try:
                os.remove(backup_path)
            except OSError:
                pass


# ==================================================
# 적용 결과 검증
# ==================================================
def verify_changes(
    file_path,
    extension,
    target_creation,
    target_access,
    target_modified,
    pdf_creation_text,
    pdf_modified_text
):
    errors = []
    tolerance = 2.0

    file_stat = os.stat(file_path)

    if row_widgets["fs_creation"]["selected"].get():
        actual = get_creation_timestamp(file_stat)

        if abs(actual - target_creation.timestamp()) > tolerance:
            errors.append("파일 만든 날짜")

    if row_widgets["fs_access"]["selected"].get():
        if abs(
            file_stat.st_atime - target_access.timestamp()
        ) > tolerance:
            errors.append("파일 액세스 날짜")

    if row_widgets["fs_modified"]["selected"].get():
        if abs(
            file_stat.st_mtime - target_modified.timestamp()
        ) > tolerance:
            errors.append("파일 수정 날짜")

    if extension == ".pdf":
        with open(file_path, "rb") as pdf_file:
            reader = PdfReader(pdf_file)
            metadata = reader.metadata or {}

            if pdf_creation_text is not None:
                expected = f"D:{pdf_creation_text}+09'00'"
                actual = str(metadata.get("/CreationDate", ""))

                if actual != expected:
                    errors.append("PDF 내부 만든 날짜")

            if pdf_modified_text is not None:
                expected = f"D:{pdf_modified_text}+09'00'"
                actual = str(metadata.get("/ModDate", ""))

                if actual != expected:
                    errors.append("PDF 내부 수정 날짜")

    return errors


# ==================================================
# 변경 실행
# ==================================================
def apply_selected_dates():
    file_path = entry_file.get().strip()

    if not os.path.isfile(file_path):
        messagebox.showerror(
            "오류",
            "올바른 파일을 선택해 주세요."
        )
        return

    if sys.platform != "win32":
        messagebox.showerror(
            "오류",
            "파일 날짜 변경은 Windows에서만 지원합니다."
        )
        return

    extension = os.path.splitext(file_path)[1].lower()

    if extension not in SUPPORTED_EXTENSIONS:
        messagebox.showerror(
            "오류",
            "PDF, TXT, CSV 파일만 변경할 수 있습니다."
        )
        return

    selected_keys = [
        key
        for key, row in row_widgets.items()
        if row["selected"].get()
    ]

    if not selected_keys:
        messagebox.showwarning(
            "선택 필요",
            "변경할 항목을 하나 이상 선택해 주세요."
        )
        return

    try:
        parsed_dates = {}

        for key in selected_keys:
            date_text = row_widgets[key]["value"].get().strip()
            parsed_dates[key] = parse_input_date(date_text)

        original_stat = os.stat(file_path)

        original_creation = datetime.fromtimestamp(
            get_creation_timestamp(original_stat)
        )

        original_access = datetime.fromtimestamp(
            original_stat.st_atime
        )

        original_modified = datetime.fromtimestamp(
            original_stat.st_mtime
        )

        target_creation = parsed_dates.get(
            "fs_creation",
            original_creation
        )

        target_access = parsed_dates.get(
            "fs_access",
            original_access
        )

        target_modified = parsed_dates.get(
            "fs_modified",
            original_modified
        )

        pdf_creation_text = (
            row_widgets["pdf_creation"]["value"].get().strip()
            if "pdf_creation" in selected_keys
            else None
        )

        pdf_modified_text = (
            row_widgets["pdf_modified"]["value"].get().strip()
            if "pdf_modified" in selected_keys
            else None
        )

        item_names = {
            "fs_creation": "파일 만든 날짜",
            "fs_access": "파일 액세스 날짜",
            "fs_modified": "파일 수정 날짜",
            "pdf_creation": "PDF 내부 만든 날짜",
            "pdf_modified": "PDF 내부 수정 날짜",
        }

        confirmation_lines = []

        for key in selected_keys:
            confirmation_lines.append(
                f"• {item_names[key]}: "
                f"{row_widgets[key]['value'].get().strip()}"
            )

        confirmed = messagebox.askyesno(
            "원본 파일 수정 확인",
            "원본 파일을 직접 수정합니다.\n"
            "별도의 결과 파일은 생성되지 않습니다.\n\n"
            f"대상 파일:\n{file_path}\n\n"
            "변경 내용:\n"
            + "\n".join(confirmation_lines)
            + "\n\n계속하시겠습니까?"
        )

        if not confirmed:
            return

        pdf_selected = (
            pdf_creation_text is not None
            or pdf_modified_text is not None
        )

        if extension == ".pdf" and pdf_selected:
            modify_pdf_in_place(
                file_path,
                pdf_creation_text,
                pdf_modified_text,
                original_creation,
                original_access,
                original_modified,
                target_creation,
                target_access,
                target_modified
            )
        else:
            set_windows_file_times(
                file_path,
                target_creation,
                target_access,
                target_modified
            )

        verification_errors = verify_changes(
            file_path,
            extension,
            target_creation,
            target_access,
            target_modified,
            pdf_creation_text,
            pdf_modified_text
        )

        # 검증 과정에서 액세스 시간이 바뀔 수 있으므로 다시 적용
        set_windows_file_times(
            file_path,
            target_creation,
            target_access,
            target_modified
        )

        load_file_information(file_path)

        if verification_errors:
            messagebox.showwarning(
                "적용 결과 확인 필요",
                "일부 항목의 적용 여부를 확인하지 못했습니다.\n\n"
                + "\n".join(
                    f"• {item}" for item in verification_errors
                )
            )
        else:
            messagebox.showinfo(
                "적용 완료",
                "선택한 날짜가 원본 파일에 적용되었습니다.\n\n"
                + "\n".join(confirmation_lines)
            )

    except PermissionError:
        messagebox.showerror(
            "적용 실패",
            "파일이 다른 프로그램에서 열려 있거나 "
            "수정 권한이 없습니다.\n\n"
            "파일을 닫은 후 다시 시도해 주세요."
        )

    except Exception as error:
        messagebox.showerror(
            "적용 실패",
            f"날짜 적용 중 오류가 발생했습니다.\n\n{error}"
        )


# ==================================================
# GUI
# ==================================================
root = tk.Tk()
root.title("PDF · TXT · CSV 날짜 변경기")
root.geometry("960x650")
root.minsize(960, 650)
root.resizable(True, True)


# 제목
tk.Label(
    root,
    text="파일 날짜 개별 변경기",
    font=("맑은 고딕", 16, "bold"),
    fg="#333333"
).pack(pady=(15, 8))


# 파일 선택
frame_file = tk.LabelFrame(
    root,
    text="파일 선택",
    padx=12,
    pady=12
)
frame_file.pack(fill="x", padx=20, pady=5)

entry_file = tk.Entry(frame_file)
entry_file.pack(
    side=tk.LEFT,
    fill="x",
    expand=True,
    padx=(0, 8)
)

tk.Button(
    frame_file,
    text="찾아보기",
    command=select_file,
    width=10
).pack(side=tk.LEFT)


# 파일 형식
frame_type = tk.Frame(root)
frame_type.pack(fill="x", padx=25, pady=(5, 0))

tk.Label(
    frame_type,
    text="파일 형식:",
    font=("맑은 고딕", 9, "bold")
).pack(side=tk.LEFT)

lbl_file_type = tk.Label(
    frame_type,
    text="-",
    fg="gray",
    font=("맑은 고딕", 9, "bold")
)
lbl_file_type.pack(side=tk.LEFT, padx=8)


# 날짜 정보 표
frame_dates = tk.LabelFrame(
    root,
    text="현재 날짜 정보 및 변경 설정",
    padx=12,
    pady=10
)
frame_dates.pack(fill="x", padx=20, pady=8)

frame_dates.columnconfigure(1, weight=1)

headers = [
    ("항목", 0),
    ("현재 날짜", 1),
    ("변경 선택", 2),
    ("변경할 날짜", 3),
    ("입력 형식", 4),
]

for text, column in headers:
    tk.Label(
        frame_dates,
        text=text,
        font=("맑은 고딕", 9, "bold"),
        fg="#333333"
    ).grid(
        row=0,
        column=column,
        padx=8,
        pady=(0, 8),
        sticky="w"
    )


row_definitions = [
    ("fs_creation", "파일 만든 날짜"),
    ("fs_access", "파일 액세스 날짜"),
    ("fs_modified", "파일 수정 날짜"),
    ("pdf_creation", "PDF 내부 만든 날짜"),
    ("pdf_modified", "PDF 내부 수정 날짜"),
]

for row_number, (key, title) in enumerate(
    row_definitions,
    start=1
):
    selected_variable = tk.BooleanVar(value=False)
    date_variable = tk.StringVar(value="")

    tk.Label(
        frame_dates,
        text=title,
        width=21,
        anchor="w",
        font=("맑은 고딕", 9, "bold")
    ).grid(
        row=row_number,
        column=0,
        padx=8,
        pady=7,
        sticky="w"
    )

    current_label = tk.Label(
        frame_dates,
        text="-",
        width=27,
        anchor="w",
        fg="gray"
    )
    current_label.grid(
        row=row_number,
        column=1,
        padx=8,
        pady=7,
        sticky="w"
    )

    check_button = tk.Checkbutton(
        frame_dates,
        text="변경",
        variable=selected_variable,
        state="disabled"
    )
    check_button.grid(
        row=row_number,
        column=2,
        padx=12,
        pady=7
    )

    date_entry = tk.Entry(
        frame_dates,
        width=20,
        textvariable=date_variable,
        font=("Consolas", 10),
        state="disabled"
    )
    date_entry.grid(
        row=row_number,
        column=3,
        padx=8,
        pady=7
    )

    tk.Label(
        frame_dates,
        text="YYYYMMDDHHMMSS",
        fg="gray"
    ).grid(
        row=row_number,
        column=4,
        padx=8,
        pady=7,
        sticky="w"
    )

    row_widgets[key] = {
        "current": current_label,
        "selected": selected_variable,
        "value": date_variable,
        "check": check_button,
        "entry": date_entry,
    }


# 전체 선택 및 해제
frame_selection = tk.Frame(root)
frame_selection.pack(pady=5)

tk.Button(
    frame_selection,
    text="전체 선택",
    command=select_all_available,
    width=12
).pack(side=tk.LEFT, padx=5)

tk.Button(
    frame_selection,
    text="전체 해제",
    command=clear_all_selections,
    width=12
).pack(side=tk.LEFT, padx=5)


# 안내
tk.Label(
    root,
    text=(
        "변경할 항목을 체크한 후 날짜를 입력하세요. "
        "입력란에는 기존 날짜가 기본값으로 표시됩니다.\n"
        "주의: 별도의 결과 파일을 만들지 않고 원본 파일을 직접 수정합니다."
    ),
    fg="#C62828",
    font=("맑은 고딕", 9),
    justify="center"
).pack(pady=(8, 5))


# 변경 실행 버튼
frame_execute = tk.Frame(
    root,
    bg="#E8F5E9",
    padx=15,
    pady=12
)
frame_execute.pack(
    fill="x",
    padx=20,
    pady=(5, 15)
)

tk.Button(
    frame_execute,
    text="변경 실행 — 선택한 날짜를 원본에 적용",
    command=apply_selected_dates,
    bg="#2E7D32",
    fg="white",
    activebackground="#1B5E20",
    activeforeground="white",
    font=("맑은 고딕", 12, "bold"),
    width=38,
    pady=9,
    cursor="hand2"
).pack()


clear_information()
root.mainloop()
