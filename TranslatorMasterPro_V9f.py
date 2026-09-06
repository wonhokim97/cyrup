# -*- coding: utf-8 -*-
"""
다중 엔진 문서 번역기 프로그램 (Final++ / Integrated) - 요청사항 반영 단일 파일

요청 반영 사항:
1) 설정 메뉴에서 'PDF 저장 프리셋' 항목 삭제
2) 1)과 관련된 코드 모두 삭제(프리셋/변수/설정/함수)
3) 문서 저장 상세 설정 창의 '한국 소설책 프리셋 적용' 버튼 제거
4) PDF 저장은 항상 소설책 보기 형식 + 양쪽 정렬(언어 무관 동일) 적용
5) 전체 종합본을 단일 파일로 제공
6) 위 요청사항 외 기존 메뉴 형태 및 기능은 그대로 유지
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont

from tkinterdnd2 import TkinterDnD, DND_FILES  # pip install tkinterdnd2

import urllib.request
import urllib.parse
import urllib.error
import json
import re
import os
import ssl
import time
import html as html_mod
import threading
import queue
import math

# ✅ fontTools(CJK 폰트 subset) 정보 로그 숨기기
import logging
logging.getLogger("fontTools").setLevel(logging.ERROR)

from bs4 import BeautifulSoup  # pip install beautifulsoup4
try:
    import pymupdf as fitz  # PyMuPDF: pip install pymupdf
except ImportError:
    import fitz  # fallback for older environments

# MS Office 문서 처리
import docx  # pip install python-docx
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from pptx import Presentation  # pip install python-pptx
import pptx.util
import openpyxl  # pip install openpyxl

SETTINGS_FILE = "translator_settings.json"

LANGUAGES = {
    "한국어": "ko", "영어": "en", "일본어": "ja",
    "중국어(간체)": "zh-CN", "중국어(번체)": "zh-TW",
    "독일어": "de", "프랑스어": "fr", "스페인어": "es", "러시아어": "ru"
}

ALLOWED_FONT_SIZES = [10, 11, 12, 14, 16, 18, 20]
ALLOWED_CHUNK_SIZES = [300, 400, 500, 600, 700, 800]

ALLOWED_ENGINES = [
    "Google 번역 (기본)", "Google 번역 (대체)", "Google (공식 API)",
    "MS 브라우저 번역", "MS (공식 API)", "MyMemory (무료)",
    "OpenAI (공식 API)", "DeepL (공식 API)"
]

ALLOWED_PAPER_SIZES = [
    "A4 (210x297mm)",
    "A5 (148x210mm)",
    "B5 (182x257mm)",
    "B6 (128x182mm)",
    "국판 (152x225mm)",
    "46판 (128x188mm)"
]

ALLOWED_DOC_FONTS = ["맑은 고딕", "바탕", "돋움", "굴림", "궁서", "Arial", "Times New Roman"]


class OverloadException(Exception):
    pass


# ----------------- 설정 로드/저장 -----------------
def load_settings():
    default_settings = {
        "target_lang": "한국어",
        "font_size": 11,
        "chunk_size": 400,
        "engine": "MS 브라우저 번역",
        "paper_size": "A4 (210x297mm)",
        "openai_api_key": "",
        "deepl_api_key": "",
        "google_api_key": "",
        "ms_api_key": "",
        "doc_font_family": "맑은 고딕",
        "doc_font_size": 11,
        "doc_line_spacing": 1.4,
        "doc_para_spacing": 10,
        "doc_margin_x": 15,
        "doc_margin_y": 20,
        "doc_page_numbers": True,
        "doc_indent_pt": 14,
        "last_url": ""
    }

    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)

            # 누락 키 채우기
            for key in default_settings:
                if key not in saved:
                    saved[key] = default_settings[key]

            # 엔진 방어
            if saved["engine"] not in ALLOWED_ENGINES:
                saved["engine"] = default_settings["engine"]

            # 용지 방어
            current_paper = str(saved.get("paper_size", "")).strip()
            if current_paper not in ALLOWED_PAPER_SIZES:
                mapped_paper = default_settings["paper_size"]
                for allowed in ALLOWED_PAPER_SIZES:
                    if current_paper and current_paper in allowed:
                        mapped_paper = allowed
                        break
                saved["paper_size"] = mapped_paper

            # 글꼴 방어
            if saved.get("doc_font_family") not in ALLOWED_DOC_FONTS:
                saved["doc_font_family"] = default_settings["doc_font_family"]

            # 들여쓰기 방어
            try:
                saved["doc_indent_pt"] = int(float(saved.get("doc_indent_pt", default_settings["doc_indent_pt"])))
            except Exception:
                saved["doc_indent_pt"] = default_settings["doc_indent_pt"]

            # (요청사항) PDF 프리셋 관련 설정이 예전 설정파일에 남아있어도 무시:
            if "pdf_preset" in saved:
                saved.pop("pdf_preset", None)

            return saved
        except Exception:
            return default_settings

    return default_settings


def save_current_config():
    settings = {
        "target_lang": lang_combobox.get() if "lang_combobox" in globals() else current_settings.get("target_lang", "한국어"),
        "font_size": font_size_var.get() if "font_size_var" in globals() else current_settings.get("font_size", 11),
        "chunk_size": chunk_size_var.get() if "chunk_size_var" in globals() else current_settings.get("chunk_size", 400),
        "engine": engine_var.get() if "engine_var" in globals() else current_settings.get("engine", "MS 브라우저 번역"),
        "paper_size": paper_size_var.get() if "paper_size_var" in globals() else current_settings.get("paper_size", "A4 (210x297mm)"),
        "openai_api_key": current_settings.get("openai_api_key", ""),
        "deepl_api_key": current_settings.get("deepl_api_key", ""),
        "google_api_key": current_settings.get("google_api_key", ""),
        "ms_api_key": current_settings.get("ms_api_key", ""),
        "doc_font_family": doc_font_family_var.get() if "doc_font_family_var" in globals() else current_settings.get("doc_font_family", "맑은 고딕"),
        "doc_font_size": doc_font_size_var.get() if "doc_font_size_var" in globals() else current_settings.get("doc_font_size", 11),
        "doc_line_spacing": doc_line_spacing_var.get() if "doc_line_spacing_var" in globals() else current_settings.get("doc_line_spacing", 1.4),
        "doc_para_spacing": doc_para_spacing_var.get() if "doc_para_spacing_var" in globals() else current_settings.get("doc_para_spacing", 10),
        "doc_margin_x": doc_margin_x_var.get() if "doc_margin_x_var" in globals() else current_settings.get("doc_margin_x", 15),
        "doc_margin_y": doc_margin_y_var.get() if "doc_margin_y_var" in globals() else current_settings.get("doc_margin_y", 20),
        "doc_page_numbers": doc_page_numbers_var.get() if "doc_page_numbers_var" in globals() else current_settings.get("doc_page_numbers", True),
        "doc_indent_pt": doc_indent_pt_var.get() if "doc_indent_pt_var" in globals() else current_settings.get("doc_indent_pt", 14),
        "last_url": current_settings.get("last_url", "")
    }

    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
    except Exception:
        pass


def save_settings():
    save_current_config()
    root.destroy()


# ----------------- SSL(보안+호환) -----------------
def urlopen_safe(req, timeout=30):
    ctx_verified = ssl.create_default_context()
    try:
        return urllib.request.urlopen(req, context=ctx_verified, timeout=timeout)
    except ssl.SSLError:
        ctx_unverified = ssl._create_unverified_context()
        return urllib.request.urlopen(req, context=ctx_unverified, timeout=timeout)


# ----------------- 상태바 테마(색/아이콘/자동복귀) -----------------
STATUS_THEMES = {
    "idle": {"bg": "#f2f2f2", "fg": "#111111", "icon": "ℹ️"},
    "work": {"bg": "#e8f4ff", "fg": "#0b3d91", "icon": "⏳"},
    "ok":   {"bg": "#e9f7ef", "fg": "#0b6623", "icon": "✅"},
    "save": {"bg": "#e6ffed", "fg": "#0b6623", "icon": "💾"},
    "warn": {"bg": "#fff8e1", "fg": "#7a4e00", "icon": "⚠️"},
    "error":{"bg": "#ffecec", "fg": "#8b0000", "icon": "❌"},
}

_status_revert_after_id = None
_current_status_level = "idle"


def _apply_status_theme(level):
    global _current_status_level
    _current_status_level = level
    th = STATUS_THEMES.get(level, STATUS_THEMES["idle"])
    status_frame.configure(bg=th["bg"])
    status_icon_label.configure(bg=th["bg"], fg=th["fg"])
    status_label.configure(bg=th["bg"], fg=th["fg"])
    metrics_label.configure(bg=th["bg"], fg=th["fg"])
    progress_label.configure(bg=th["bg"], fg=th["fg"])
    progress_right.configure(bg=th["bg"])
    status_icon_var.set(th["icon"])


def set_status(level, text, auto_revert_seconds=None):
    """
    level: idle/work/ok/save/warn/error
    auto_revert_seconds: 지정 시, 해당 시간 후 idle로 자동 복귀(단, 번역중이면 work 유지)
    """
    global _status_revert_after_id
    status_var.set(text)
    _apply_status_theme(level)

    if _status_revert_after_id:
        try:
            root.after_cancel(_status_revert_after_id)
        except Exception:
            pass
        _status_revert_after_id = None

    if auto_revert_seconds is not None:
        def _revert():
            if is_translating:
                set_status("work", status_var.get())
            else:
                set_status("idle", "상태: 대기 중")
        _status_revert_after_id = root.after(int(auto_revert_seconds * 1000), _revert)


def format_eta(seconds):
    if seconds is None:
        return "--:--"
    seconds = max(0, int(seconds))
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h > 0:
        return f"{h:d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


# ----------------- PDF 저장 폰트 탐색 -----------------
def _find_font_path_for_pdf(font_family):
    win_map = {
        "맑은 고딕": ["malgun.ttf", "MalgunGothic.ttf"],
        "바탕": ["batang.ttc", "Batang.ttc"],
        "돋움": ["dotum.ttc", "Dotum.ttc"],
        "굴림": ["gulim.ttc", "Gulim.ttc"],
        "궁서": ["gungsuh.ttc", "Gungsuh.ttc"],
        "Arial": ["arial.ttf", "Arial.ttf"],
        "Times New Roman": ["times.ttf", "Times.ttf", "Times New Roman.ttf"]
    }

    nix_candidates = {
        "맑은 고딕": [
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "/Library/Fonts/AppleGothic.ttf",
        ],
        "바탕": [
            "/usr/share/fonts/truetype/noto/NotoSerifCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
        ],
        "Arial": [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/Library/Fonts/Arial.ttf",
        ],
        "Times New Roman": [
            "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
            "/Library/Fonts/Times New Roman.ttf",
        ],
    }

    if os.name == "nt":
        base = r"C:\Windows\Fonts"
        for fn in win_map.get(font_family, []):
            p = os.path.join(base, fn)
            if os.path.exists(p):
                return p
        p = os.path.join(base, "malgun.ttf")
        if os.path.exists(p):
            return p
        return None

    for p in nix_candidates.get(font_family, []):
        if os.path.exists(p):
            return p

    common = [
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/truetype/noto/NotoSerifCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/AppleGothic.ttf",
    ]
    for p in common:
        if os.path.exists(p):
            return p
    return None


# ----------------- 번역 엔진 -----------------
EDGE_TOKEN = None
EDGE_TOKEN_TIME = 0


def translate_openai(text, target_lang_code, settings_snapshot):
    api_key = settings_snapshot.get("openai_api_key", "").strip()
    if not api_key:
        return "\n[OpenAI API 키가 설정되지 않았습니다.]\n"

    lang_name = "Korean"
    for k, v in LANGUAGES.items():
        if v == target_lang_code:
            lang_name = k
            break

    payload = {
        "model": "gpt-4o-mini",
        "messages": [
            {"role": "system", "content": f"Translate to {lang_name}. Output ONLY the translation."},
            {"role": "user", "content": text}
        ],
        "temperature": 0.3
    }

    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    )

    try:
        with urlopen_safe(req) as response:
            return json.loads(response.read().decode("utf-8"))["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"\n[OpenAI 오류: {str(e)}]\n"


def translate_deepl(text, target_lang_code, settings_snapshot):
    api_key = settings_snapshot.get("deepl_api_key", "").strip()
    if not api_key:
        return "\n[DeepL API 키가 설정되지 않았습니다.]\n"

    dl_map = {"ko": "KO", "en": "EN-US", "ja": "JA", "zh-CN": "ZH", "zh-TW": "ZH"}
    dl_target = dl_map.get(target_lang_code, target_lang_code.upper())

    domain = "api-free.deepl.com" if api_key.endswith(":fx") else "api.deepl.com"
    data = urllib.parse.urlencode({"text": text, "target_lang": dl_target}).encode("utf-8")

    req = urllib.request.Request(
        f"https://{domain}/v2/translate",
        data=data,
        headers={"Authorization": f"DeepL-Auth-Key {api_key}", "Content-Type": "application/x-www-form-urlencoded"}
    )

    try:
        with urlopen_safe(req) as response:
            return json.loads(response.read().decode("utf-8"))["translations"][0]["text"]
    except Exception as e:
        return f"\n[DeepL 오류: {str(e)}]\n"


def translate_google_official(text, target_lang_code, settings_snapshot):
    api_key = settings_snapshot.get("google_api_key", "").strip()
    if not api_key:
        return "\n[Google 공식 API 키가 설정되지 않았습니다.]\n"

    url = f"https://translation.googleapis.com/language/translate/v2?key={api_key}"
    req = urllib.request.Request(
        url,
        data=json.dumps({"q": text, "target": target_lang_code}).encode("utf-8"),
        headers={"Content-Type": "application/json"}
    )

    try:
        with urlopen_safe(req) as response:
            parsed = json.loads(response.read().decode("utf-8"))
            return html_mod.unescape(parsed["data"]["translations"][0]["translatedText"])
    except Exception as e:
        return f"\n[Google 공식 오류: {str(e)}]\n"


def translate_ms_official(text, target_lang_code, settings_snapshot):
    api_key_raw = settings_snapshot.get("ms_api_key", "").strip()
    if not api_key_raw:
        return "\n[MS 공식 API 키가 설정되지 않았습니다.]\n"

    parts = api_key_raw.split(",")
    api_key = parts[0].strip()
    region = parts[1].strip() if len(parts) > 1 else ""

    url = f"https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&to={target_lang_code}"
    headers = {"Ocp-Apim-Subscription-Key": api_key, "Content-Type": "application/json"}
    if region:
        headers["Ocp-Apim-Subscription-Region"] = region

    req = urllib.request.Request(url, data=json.dumps([{"Text": text}]).encode("utf-8"), headers=headers)

    try:
        with urlopen_safe(req) as response:
            return json.loads(response.read().decode("utf-8"))[0]["translations"][0]["text"]
    except Exception as e:
        return f"\n[MS 공식 오류: {str(e)}]\n"


def translate_ms_edge(text, target_lang_code):
    global EDGE_TOKEN, EDGE_TOKEN_TIME

    for _ in range(3):
        try:
            if not EDGE_TOKEN or time.time() - EDGE_TOKEN_TIME > 600:
                req_auth = urllib.request.Request(
                    "https://edge.microsoft.com/translate/auth",
                    headers={"User-Agent": "Mozilla/5.0"}
                )
                with urlopen_safe(req_auth) as res:
                    EDGE_TOKEN = res.read().decode("utf-8")
                    EDGE_TOKEN_TIME = time.time()

            url = f"https://api-edge.cognitive.microsofttranslator.com/translate?from=&to={target_lang_code}&api-version=3.0"
            req = urllib.request.Request(
                url,
                data=json.dumps([{"Text": text}]).encode("utf-8"),
                headers={
                    "Authorization": f"Bearer {EDGE_TOKEN}",
                    "Content-Type": "application/json",
                    "User-Agent": "Mozilla/5.0"
                }
            )
            with urlopen_safe(req) as res:
                return json.loads(res.read().decode("utf-8"))[0]["translations"][0]["text"]

        except urllib.error.HTTPError as e:
            if e.code == 429:
                raise OverloadException()
            elif e.code == 401:
                EDGE_TOKEN = None
                continue
            return f"\n[MS 브라우저 번역 오류 (HTTP {e.code})]\n"
        except OverloadException:
            raise
        except Exception as e:
            return f"\n[MS 브라우저 번역 오류: {str(e)}]\n"


def translate_google(text, target_lang_code, client="gtx"):
    q = urllib.parse.quote(text)
    url = f"https://translate.googleapis.com/translate_a/single?client={client}&sl=auto&tl={target_lang_code}&dt=t&q={q}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urlopen_safe(req) as res:
            parsed = json.loads(res.read().decode("utf-8"))
            trans_text = "".join([s[0] for s in parsed[0] if s[0]])
            return re.sub(r"[a-f0-9]{32}", "", trans_text, flags=re.IGNORECASE)
    except urllib.error.HTTPError as e:
        if e.code == 429:
            raise OverloadException()
        return "\n[Google 번역 오류]\n"
    except OverloadException:
        raise
    except Exception as e:
        return f"\n[Google 번역 오류: {str(e)}]\n"


def translate_mymemory(text, target_lang_code):
    url = "https://api.mymemory.translated.net/get"
    data = urllib.parse.urlencode({"q": text, "langpair": f"Autodetect|{target_lang_code}"}).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"User-Agent": "Mozilla/5.0"})

    try:
        with urlopen_safe(req) as res:
            parsed = json.loads(res.read().decode("utf-8"))
            if parsed.get("responseStatus") == 429:
                raise OverloadException()
            return parsed.get("responseData", {}).get("translatedText", text)
    except urllib.error.HTTPError as e:
        if e.code == 429:
            raise OverloadException()
        return "\n[MyMemory 번역 오류]\n"
    except OverloadException:
        raise
    except Exception as e:
        return f"\n[MyMemory 번역 오류: {str(e)}]\n"


def translate_text(text, target_lang_code, engine_name, settings_snapshot):
    if engine_name == "OpenAI (공식 API)":
        return translate_openai(text, target_lang_code, settings_snapshot)
    elif engine_name == "DeepL (공식 API)":
        return translate_deepl(text, target_lang_code, settings_snapshot)
    elif engine_name == "Google (공식 API)":
        return translate_google_official(text, target_lang_code, settings_snapshot)
    elif engine_name == "MS (공식 API)":
        return translate_ms_official(text, target_lang_code, settings_snapshot)
    elif engine_name == "MS 브라우저 번역":
        return translate_ms_edge(text, target_lang_code)
    elif engine_name == "MyMemory (무료)":
        return translate_mymemory(text, target_lang_code)
    elif engine_name == "Google 번역 (대체)":
        return translate_google(text, target_lang_code, "dict-chrome-ex")
    else:
        return translate_google(text, target_lang_code, "gtx")


# ----------------- 번역 스레드/진행률 -----------------
translate_thread = None
translate_queue = queue.Queue()
translate_stop_event = threading.Event()
is_translating = False

# 번역 통계
stats_start_time = None
stats_total_blocks = 0
stats_done_blocks = 0
stats_total_chars = 0
stats_done_chars = 0


def _set_progress_percent(pct):
    pct = max(0, min(100, int(pct)))
    progress_var.set(pct)
    progress_label_var.set(f"{pct}%")
    progressbar.update_idletasks()


def _ui_set_running(running: bool):
    global is_translating
    is_translating = running
    try:
        btn_translate.config(state=tk.DISABLED if running else tk.NORMAL)
        btn_cancel.config(state=tk.NORMAL if running else tk.DISABLED)
    except Exception:
        pass


def _update_metrics():
    if not stats_start_time or stats_total_blocks <= 0:
        metrics_var.set("")
        return

    elapsed = time.time() - stats_start_time
    elapsed = max(0.001, elapsed)

    blocks_per_sec = stats_done_blocks / elapsed
    blocks_per_min = blocks_per_sec * 60
    chars_per_sec = stats_done_chars / elapsed

    remaining_blocks = max(0, stats_total_blocks - stats_done_blocks)
    eta_sec = (remaining_blocks / blocks_per_sec) if blocks_per_sec > 0 else None

    metrics_var.set(
        f"블록 {stats_done_blocks}/{stats_total_blocks}  |  평균 {blocks_per_min:.1f} blk/min, {chars_per_sec:.0f} ch/s  |  소요 {format_eta(elapsed)}  |  남은 {format_eta(eta_sec)}"
    )


def on_cancel_translate():
    if is_translating:
        translate_stop_event.set()
        set_status("warn", "상태: 중단 요청됨... (현재 블록 처리 후 중단)", auto_revert_seconds=None)


def _poll_translate_queue():
    try:
        while True:
            msg = translate_queue.get_nowait()
            mtype = msg[0]

            if mtype == "clear_output":
                text_output.delete("1.0", tk.END)

            elif mtype == "append_output":
                text_output.insert(tk.END, msg[1])
                text_output.see(tk.END)

            elif mtype == "status":
                set_status("work" if is_translating else "idle", msg[1])

            elif mtype == "stats":
                global stats_done_blocks, stats_done_chars, stats_total_blocks, stats_total_chars, stats_start_time
                stats_done_blocks, stats_done_chars, stats_total_blocks, stats_total_chars, stats_start_time = msg[1]
                _update_metrics()

                if stats_total_blocks > 0:
                    pct = (stats_done_blocks / stats_total_blocks) * 100
                    _set_progress_percent(pct)

            elif mtype == "done":
                _ui_set_running(False)
                _set_progress_percent(100)
                _update_metrics()
                set_status("ok", "상태: 번역 완료.", auto_revert_seconds=4)
                messagebox.showinfo("완료", msg[1] if len(msg) > 1 else "번역이 완료되었습니다.")

            elif mtype == "cancelled":
                _ui_set_running(False)
                _update_metrics()
                set_status("warn", "상태: 번역이 중단되었습니다.", auto_revert_seconds=4)
                messagebox.showinfo("중단", "사용자 요청으로 번역이 중단되었습니다.")

            elif mtype == "overload":
                _ui_set_running(False)
                text_output.insert(tk.END, "\n\n[강제 중단] 서버 과부하 감지. 번역을 즉시 중단했습니다.")
                set_status("error", "상태: 서버 과부하 감지, 번역 즉시 중단됨", auto_revert_seconds=6)
                messagebox.showwarning("강제 중단", "번역 서버 과부하가 감지되었습니다. 잠시 후 시도하세요.")

            elif mtype == "error":
                _ui_set_running(False)
                set_status("error", "상태: 번역 중 오류 발생", auto_revert_seconds=6)
                messagebox.showerror("오류", msg[1])

            translate_queue.task_done()

    except queue.Empty:
        pass

    root.after(60, _poll_translate_queue)


def _translate_worker(source_text, target_lang_code, engine_name, chunk_limit, settings_snapshot):
    global stats_start_time, stats_total_blocks, stats_done_blocks, stats_total_chars, stats_done_chars

    try:
        blocks = []
        current_batch = []
        current_len = 0
        SAFE_LIMIT = int(chunk_limit)

        all_lines = source_text.split("\n")
        total_chars = len(source_text)

        for line in all_lines:
            if translate_stop_event.is_set():
                translate_queue.put(("cancelled",))
                return

            if current_len + len(line) > SAFE_LIMIT and current_batch:
                blocks.append(current_batch)
                current_batch = [line]
                current_len = len(line)
            else:
                current_batch.append(line)
                current_len += len(line) + 1

        if current_batch:
            blocks.append(current_batch)

        total_blocks = len(blocks)

        stats_start_time = time.time()
        stats_total_blocks = max(1, total_blocks)
        stats_done_blocks = 0
        stats_total_chars = max(1, total_chars)
        stats_done_chars = 0

        translate_queue.put(("clear_output",))
        translate_queue.put(("status", f"상태:[{engine_name}] {SAFE_LIMIT}자 단위 번역 시작..."))
        translate_queue.put(("stats", (stats_done_blocks, stats_done_chars, stats_total_blocks, stats_total_chars, stats_start_time)))

        def call_translate_with_backoff(text_to_translate):
            backoff = 0.8
            for _attempt in range(3):
                if translate_stop_event.is_set():
                    return None
                try:
                    return translate_text(text_to_translate, target_lang_code, engine_name, settings_snapshot)
                except OverloadException:
                    raise
                except Exception:
                    time.sleep(backoff)
                    backoff *= 2
            return translate_text(text_to_translate, target_lang_code, engine_name, settings_snapshot)

        for i, batch in enumerate(blocks, start=1):
            if translate_stop_event.is_set():
                translate_queue.put(("cancelled",))
                return

            translate_queue.put(("status", f"상태: 번역 진행 중... ({i} / {total_blocks}) 블록 처리 중"))

            block_text = "\n".join(batch)
            block_chars = len(block_text)

            if not block_text.strip():
                keep_newlines = "\n" * (len(batch) if i < total_blocks else max(len(batch) - 1, 0))
                translate_queue.put(("append_output", keep_newlines))
                stats_done_blocks += 1
                stats_done_chars += block_chars
                translate_queue.put(("stats", (stats_done_blocks, stats_done_chars, stats_total_blocks, stats_total_chars, stats_start_time)))
                continue

            res = call_translate_with_backoff(block_text)
            if res is None:
                translate_queue.put(("cancelled",))
                return

            res_lines = res.split("\n")
            if len(res_lines) == len(batch):
                batch_result = "\n".join(res_lines)
            else:
                recovered = []
                for line in batch:
                    if translate_stop_event.is_set():
                        translate_queue.put(("cancelled",))
                        return

                    if not line.strip():
                        recovered.append("")
                    else:
                        one = call_translate_with_backoff(line)
                        if one is None:
                            translate_queue.put(("cancelled",))
                            return
                        recovered.append(one.replace("\n", " "))
                        time.sleep(0.2)
                batch_result = "\n".join(recovered)

            translate_queue.put(("append_output", batch_result + ("\n" if i < total_blocks else "")))

            stats_done_blocks += 1
            stats_done_chars += block_chars
            translate_queue.put(("stats", (stats_done_blocks, stats_done_chars, stats_total_blocks, stats_total_chars, stats_start_time)))

            if i < total_blocks:
                time.sleep(0.35)

        translate_queue.put(("done", f"[{engine_name}] 번역이 완료되었습니다."))

    except OverloadException:
        translate_queue.put(("overload",))
    except Exception as e:
        translate_queue.put(("error", f"번역 중 오류가 발생했습니다.\n{str(e)}"))


def on_translate():
    global translate_thread
    if is_translating:
        return

    source_text = text_input.get("1.0", "end-1c")
    if not source_text.strip():
        messagebox.showwarning("알림", "번역할 텍스트를 입력하세요.")
        return

    engine_name = engine_var.get()
    target_lang_code = LANGUAGES.get(lang_combobox.get(), "ko")
    chunk_limit = chunk_size_var.get()

    settings_snapshot = dict(current_settings)
    settings_snapshot["engine"] = engine_name

    translate_stop_event.clear()
    _ui_set_running(True)

    _set_progress_percent(0)
    metrics_var.set("")
    set_status("work", "상태: 번역 준비 중...", auto_revert_seconds=None)

    translate_thread = threading.Thread(
        target=_translate_worker,
        args=(source_text, target_lang_code, engine_name, chunk_limit, settings_snapshot),
        daemon=True
    )
    translate_thread.start()


# ----------------- 문서 추출 -----------------
def extract_text_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".txt":
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read()
        except UnicodeDecodeError:
            with open(file_path, "r", encoding="cp949") as f:
                return f.read()

    elif ext == ".pdf":
        extracted = ""
        doc = fitz.open(file_path)
        for page in doc:
            t = page.get_text("text")
            if t:
                extracted += t + "\n\n"
        doc.close()

        pattern = r"[^a-zA-Z0-9\uAC00-\uD7A3\u3131-\u318E\u4E00-\u9FFF\u3040-\u30FF\s\.,!?\'\"\(\)\[\]\-\+=/%:;<>~]"
        cleaned = re.sub(pattern, " ", extracted)
        cleaned = re.sub(r"\n\s*\n", "<P>", cleaned)
        cleaned = cleaned.replace("\n", " ").replace("<P>", "\n\n")
        return re.sub(r" {2,}", " ", cleaned)

    elif ext == ".docx":
        return "\n".join([p.text for p in docx.Document(file_path).paragraphs if p.text.strip()])

    elif ext == ".pptx":
        extracted = []
        prs = Presentation(file_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    extracted.append(shape.text.strip())
            extracted.append("")
        return "\n".join(extracted)

    elif ext == ".xlsx":
        extracted = []
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                row_text = [str(c) for c in row if c is not None]
                if row_text:
                    extracted.append("\t".join(row_text))
            extracted.append("")
        return "\n".join(extracted)

    else:
        raise ValueError("지원하지 않는 파일 형식입니다.")


def on_reset():
    if is_translating:
        messagebox.showwarning("알림", "번역 중에는 초기화할 수 없습니다. 먼저 중단하세요.")
        return
    text_input.delete("1.0", tk.END)
    text_output.delete("1.0", tk.END)
    _set_progress_percent(0)
    metrics_var.set("")
    set_status("idle", "상태: 초기화되었습니다.", auto_revert_seconds=2)


def load_from_file():
    if is_translating:
        messagebox.showwarning("알림", "번역 중에는 파일 불러오기를 할 수 없습니다. 먼저 중단하세요.")
        return

    file_path = filedialog.askopenfilename(
        filetypes=[("지원 문서", "*.txt;*.pdf;*.docx;*.pptx;*.xlsx"), ("모든 파일", "*.*")]
    )
    if file_path:
        try:
            text = extract_text_from_file(file_path)
            text_input.delete("1.0", tk.END)
            text_input.insert(tk.END, text.strip())
            set_status("ok", f"상태: 문서 불러오기 완료 ({os.path.basename(file_path)})", auto_revert_seconds=3)
        except Exception as e:
            set_status("error", "상태: 문서 불러오기 실패", auto_revert_seconds=5)
            messagebox.showerror("오류", f"문서를 읽지 못했습니다.\n{str(e)}")


def load_from_url():
    if is_translating:
        messagebox.showwarning("알림", "번역 중에는 웹 불러오기를 할 수 없습니다. 먼저 중단하세요.")
        return

    result_url = []
    dialog = tk.Toplevel(root)
    dialog.title("웹사이트 불러오기")
    dialog.geometry("700x200")
    dialog.transient(root)
    dialog.grab_set()

    dialog.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - 350
    y = root.winfo_y() + (root.winfo_height() // 2) - 90
    dialog.geometry(f"+{x}+{y}")

    tk.Label(dialog, text="불러올 웹사이트 주소를 입력하세요:\n(예: www.wikipedia.org)", font=global_font).pack(pady=(20, 10))
    url_entry = tk.Entry(dialog, width=65, font=global_font)
    url_entry.pack(fill=tk.X, padx=30, pady=5)

    last_saved_url = current_settings.get("last_url", "")
    if last_saved_url:
        url_entry.insert(0, last_saved_url)
        url_entry.select_range(0, tk.END)

    url_entry.focus()

    def on_confirm(event=None):
        entered_url = url_entry.get().strip()
        result_url.append(entered_url)
        current_settings["last_url"] = entered_url
        save_current_config()
        dialog.destroy()

    url_entry.bind("<Return>", on_confirm)

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=15)
    tk.Button(btn_frame, text="확인", font=global_font, command=on_confirm, bg="#4CAF50", fg="white", width=10).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="취소", font=global_font, command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=10)

    dialog.wait_window()

    if not result_url or not result_url[0]:
        return

    url = result_url[0]
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    start_time = time.time()
    
    def _update_web_progress(pct):
        elapsed = time.time() - start_time
        if pct > 0 and pct < 100:
            total_est = (elapsed / pct) * 100
            eta_sec = total_est - elapsed
        else:
            eta_sec = None
        
        metrics_var.set(f"소요 {format_eta(elapsed)}  |  남은 {format_eta(eta_sec)}")
        _set_progress_percent(pct)

    _update_web_progress(10)
    set_status("work", "상태: 웹사이트를 분석하는 중입니다...", auto_revert_seconds=None)
    root.update()

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen_safe(req) as res:
            raw_data = res.read()
            _update_web_progress(40)
            try:
                html_data = raw_data.decode("utf-8")
            except UnicodeDecodeError:
                html_data = raw_data.decode("cp949", errors="ignore")

        soup = BeautifulSoup(html_data, "html.parser")
        _update_web_progress(60)

        for element in soup(["script", "style", "nav", "footer", "header", "aside", "noscript", "iframe", "form"]):
            element.extract()

        noise = re.compile(r"(ad|ads|advert|banner|sidebar|menu|popup|comment|share|social|widget|promo)\b", re.IGNORECASE)
        for el in soup.find_all(["div", "ul", "li", "section"]):
            cls = " ".join(el.get("class", []))
            eid = el.get("id", "")
            if noise.search(cls + " " + eid):
                el.extract()

        main_content = soup.find("article") or soup.find("main") or soup.find("body") or soup

        paragraphs = []

        def norm_text(s: str) -> str:
            s = s.replace("\xa0", " ")
            s = re.sub(r"[ \t]+", " ", s)
            s = re.sub(r"\n{2,}", " ", s)
            return s.strip()

        # 1️⃣ <p> 태그 기반 문단 수집
        ps = main_content.find_all("p")
        for p in ps:
            txt = norm_text(p.get_text(" ", strip=True))
            if txt:
                paragraphs.append(txt)

        _update_web_progress(80)

        # 2️⃣ <p>가 거의 없거나 없으면 → 전체 텍스트 fallback
        if len(paragraphs) < 2:
            raw = main_content.get_text(separator="\n")
            lines = [norm_text(ln) for ln in raw.split("\n") if ln.strip()]

            buf = []
            for ln in lines:
                buf.append(ln)

            if buf:
                paragraphs = buf

        _update_web_progress(90)

        # 3️⃣ 중복 제거 (순서 유지)
        deduped = []
        seen = set()
        for p in paragraphs:
            key = re.sub(r"\s+", " ", p).lower()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(p)

        cleaned_text = "\n".join(deduped)

        text_input.delete("1.0", tk.END)
        text_input.insert(tk.END, cleaned_text)
        _update_web_progress(100)
        set_status("ok", "상태: 웹사이트 텍스트 불러오기 완료.", auto_revert_seconds=3)

    except Exception as e:
        _set_progress_percent(0)
        metrics_var.set("")
        set_status("error", "상태: 웹사이트 불러오기 실패", auto_revert_seconds=6)
        messagebox.showerror("오류", f"불러오기 실패.\n{str(e)}")


def handle_drop(event):
    if is_translating:
        return
    try:
        file_list = root.tk.splitlist(event.data)
    except Exception:
        file_list = str(event.data).split()

    appended_any = False
    for file_path in file_list:
        file_path = file_path.strip()
        if not os.path.isfile(file_path):
            continue
        try:
            extracted_text = extract_text_from_file(file_path).strip()
            if not extracted_text:
                continue
            if text_input.get("1.0", tk.END).strip():
                text_input.insert(tk.END, f"\n\n---[추가됨: {os.path.basename(file_path)}] ---\n\n")
            text_input.insert(tk.END, extracted_text)
            appended_any = True
        except Exception:
            pass

    if appended_any:
        set_status("ok", "상태: 드래그앤드롭으로 문서를 추가했습니다.", auto_revert_seconds=2)


# ===============================
# 텍스트 전처리 유틸 영역
# ===============================

def refine_text_for_readability(text: str) -> str:
    """
    PDF 저장 시 가독성 개선용 문장 다듬기
    - 줄바꿈 정리
    - 과도한 행 분절 완화
    - 문단 흐름 자연화
    (의미 변화 없음)
    """
    # 줄바꿈 정규화
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # 문단 내부 줄바꿈 → 공백으로 연결
    paragraphs = re.split(r"\n\s*\n+", text.strip())
    refined_paras = []

    for p in paragraphs:
        lines = [ln.strip() for ln in p.split("\n") if ln.strip()]
        if not lines:
            continue
        refined_paras.append(" ".join(lines))

    return "\n\n".join(refined_paras)


# ----------------- 저장 로직 -----------------
def _indent_spaces_from_pt(indent_pt):
    """TXT/PPTX 등에서 '들여쓰기(pt)'를 공백 수로 근사 변환."""
    try:
        n = int(round(float(indent_pt) / 3.5))  # 14pt ≈ 4칸
    except Exception:
        n = 4
    return max(0, min(24, n))

def split_paragraphs_preserve_lines(content: str) -> list[str]:
    """
    ✅ PDF 저장용 문단 분리:
    - 빈 줄(연속 개행) 기준으로 문단 분리
    - 문단 내부의 줄바꿈(문장/라인 구조)은 그대로 유지
    """
    content = content.replace("\r\n", "\n").replace("\r", "\n").strip()
    paras = re.split(r"\n\s*\n+", content)
    return [p.rstrip() for p in paras if p.strip()]

def _split_lines_as_paragraphs(content: str):
    """줄 단위로 문단 취급 (EPUB 포함 기본 호환)"""
    raw = content.replace("\r\n", "\n").replace("\r", "\n")
    return [p.strip() for p in raw.split("\n") if p.strip()]


def _split_blanklines_as_paragraphs(content: str):
    """소설/출판형: 빈 줄 기준 문단 분리. 내부 개행은 공백으로 합쳐 자연스러운 흐름"""
    raw = content.replace("\r\n", "\n").replace("\r", "\n").strip()
    parts = re.split(r"\n\s*\n+", raw)
    out = []
    for part in parts:
        lines = [ln.strip() for ln in part.split("\n") if ln.strip()]
        if lines:
            out.append(" ".join(lines))
    return out


def _wrap_paragraph_bookwise(
    pdf, text: str, first_w: float, rest_w: float,
    min_next_words: int = 2, min_next_ratio: float = 0.55,
    max_balance_pass: int = 3
):
    """
    ✅ 소설책 조판용 줄 나눔 (음절 단위 줄바꿈 허용)
    - 공백 기반으로 단어를 붙이되, 줄 끝에 단어가 다 안 들어가면 글자 단위로 잘라서 채움.
    - 양쪽 정렬 시 단어 간격(띄어쓰기)이 과도하게 벌어지는 현상을 방지함.
    """
    import re, html

    # 문단에 한글이 포함되어 있는지 확인하여 한국어 여부 동적 판별
    is_korean = bool(re.search(r'[가-힣]', text))

    text = html.unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return []

    tokens = text.split(" ")

    def width(s: str) -> float:
        return pdf.get_string_width(s)

    def limit(idx: int) -> float:
        return first_w if idx == 0 else rest_w

    lines = []
    i = 0
    while i < len(tokens):
        wlim = limit(len(lines))
        cur = ""
        
        while i < len(tokens):
            nxt = tokens[i]
            if not nxt:
                i += 1
                continue
                
            cand = nxt if not cur else cur + " " + nxt
            if width(cand) <= wlim:
                # 단어가 통째로 들어가는 경우
                cur = cand
                i += 1
            else:
                if not is_korean and cur:
                    # 단어 단위 줄바꿈: 현재 줄에 더 이상 단어가 들어가지 않으면 다음 줄로 넘김
                    break

                # 단어가 통째로 안 들어가면, 들어갈 수 있는 글자까지만 자르고 나머지는 다음 줄로 넘김 (음절 단위 줄바꿈)
                partial = ""
                remains = ""
                prefix = cur + " " if cur else ""
                
                for j, ch in enumerate(nxt):
                    if width(prefix + partial + ch) <= wlim:
                        partial += ch
                    else:
                        remains = nxt[j:]
                        break
                
                if partial:
                    cur = prefix + partial
                    tokens[i] = remains
                elif not cur:
                    # 첫 글자조차 안 들어가는 극단적인 경우 (글자 하나가 한 줄보다 넓음)
                    cur = nxt[0]
                    if len(nxt) > 1:
                        tokens[i] = nxt[1:]
                    else:
                        i += 1
                break
                
        if cur:
            lines.append(cur)

    return lines


def _pdf_ensure_space(pdf, required_h, page_h, margin_y, margin_x):
    """페이지 분리 보장(중첩 방지)"""
    if pdf.get_y() + required_h > page_h - margin_y:
        pdf.add_page()
        pdf.set_xy(margin_x, margin_y)


from fpdf.enums import XPos, YPos

def _pdf_draw_line_justify_safe(
    pdf, line: str, x: float, y: float, width: float, line_h: float,
    justify: bool, is_last_line: bool,
    max_extra_space_mm: float = 2.0,
    max_char_spacing_pt: float = 0.8
):
    clean = (line or "").rstrip()
    if not clean:
        pdf.set_y(y + line_h)
        return

    text_width = pdf.get_string_width(clean)
    
    if (not justify) or is_last_line or text_width >= width or len(clean) < 2:
        pdf.set_xy(x, y)
        pdf.cell(width, line_h, clean, align="L")
        pdf.set_y(y + line_h)
        return

    total_extra = width - text_width
    num_spaces = clean.count(" ")
    num_chars = len(clean) - 1
    
    # PT를 mm로 변환 (1 pt = 0.352777778 mm)
    max_char_spacing_mm = max_char_spacing_pt * 0.352777778

    cx = x
    
    if num_spaces > 0:
        extra_per_space = total_extra / num_spaces
        # 소설책 형식: 띄어쓰기 간격이 너무 벌어지면 글자 간격으로 분산
        if extra_per_space > max_extra_space_mm:
            desired_extra_per_char = (total_extra - (max_extra_space_mm * num_spaces)) / max(1, num_chars)
            
            if desired_extra_per_char > max_char_spacing_mm:
                # 글자 간격이 한계치를 초과하면 최대치로 고정하고, 나머지를 다시 띄어쓰기에 몰아줌 (단어 붕괴 방지)
                extra_per_char = max_char_spacing_mm
                remaining_for_spaces = total_extra - (extra_per_char * num_chars)
                extra_per_space = remaining_for_spaces / num_spaces
            else:
                extra_per_space = max_extra_space_mm
                extra_per_char = desired_extra_per_char
        else:
            extra_per_char = 0.0

        if extra_per_char > 0:
            # 글자 간격도 늘려야 하는 경우 (단어 내부 글자씩 렌더링)
            for i, ch in enumerate(clean):
                ww = pdf.get_string_width(ch)
                if ch == " ":
                    ww += extra_per_space
                pdf.set_xy(cx, y)
                pdf.cell(ww, line_h, ch, new_x=XPos.RIGHT, new_y=YPos.TOP)
                cx += ww
                if i < len(clean) - 1:
                    cx += extra_per_char
        else:
            # 띄어쓰기 간격만 늘리는 경우 (기존 방식 유지하여 효율성 확보)
            words = clean.split(" ")
            for i, word in enumerate(words):
                if word:
                    ww = pdf.get_string_width(word)
                    pdf.set_xy(cx, y)
                    pdf.cell(ww, line_h, word, new_x=XPos.RIGHT, new_y=YPos.TOP)
                    cx += ww
                if i < len(words) - 1:
                    cx += pdf.get_string_width(" ") + extra_per_space

    else:
        # 공백이 없는 경우 (예: 중국어/일본어) - 글자 간격으로 분산
        extra_per_char = total_extra / max(1, num_chars)
        
        # 글자가 과도하게 벌어지는 것을 방지 (가독성 우선)
        if extra_per_char > max_char_spacing_mm:
            extra_per_char = max_char_spacing_mm

        for i, ch in enumerate(clean):
            ww = pdf.get_string_width(ch)
            pdf.set_xy(cx, y)
            pdf.cell(ww, line_h, ch, new_x=XPos.RIGHT, new_y=YPos.TOP)
            cx += ww
            if i < len(clean) - 1:
                cx += extra_per_char

    pdf.set_y(y + line_h)


def pdf_write_paragraph_novel(
    pdf, paragraph, page_w, page_h,
    margin_x, margin_y,
    base_indent_pt, line_spacing, para_spacing_pt
):
    # ✅ 사용자의 상세 설정(들여쓰기) 조건을 있는 그대로 적용
    if not paragraph or not paragraph.strip():
        return

    PT_TO_MM = 0.352777778
    BASE_LEADING_FACTOR = 1.25

    # 들여쓰기 & 폭 계산
    indent_mm = base_indent_pt * PT_TO_MM
    effective_w = page_w - (margin_x * 2)

    first_x = margin_x + indent_mm
    rest_x = margin_x

    first_w = effective_w - indent_mm
    rest_w = effective_w

    # 줄 높이(mm)
    line_h = (
        pdf.font_size_pt *
        BASE_LEADING_FACTOR *
        float(line_spacing)
    ) * PT_TO_MM

    lines = _wrap_paragraph_bookwise(
        pdf,
        paragraph,
        first_w=first_w,
        rest_w=rest_w
    )
    if not lines:
        return

    # 출력: 양쪽정렬(마지막 줄 좌측정렬)
    for idx, ln in enumerate(lines):
        _pdf_ensure_space(pdf, line_h, page_h, margin_y, margin_x)
        y = pdf.get_y()

        x = first_x if idx == 0 else rest_x
        w = first_w if idx == 0 else rest_w

        _pdf_draw_line_justify_safe(
            pdf,
            ln,
            x=x,
            y=y,
            width=w,
            line_h=line_h,
            justify=True,
            is_last_line=(idx == len(lines) - 1),
            max_extra_space_mm=2.0,
            max_char_spacing_pt=0.8
        )

    # 문단 간격
    gap_mm = float(para_spacing_pt) * PT_TO_MM
    _pdf_ensure_space(pdf, gap_mm + 0.1, page_h, margin_y, margin_x)
    pdf.ln(gap_mm)


def normalize_news_paragraphs_for_pdf(text: str) -> str:
    """
    뉴스 기사 전용 문단 경계 복원(PDF 저장 전용)
    - Python re 고정폭 제약 회피(look-behind 없음)
    """
    import re
    if not text:
        return ""
    t = text.replace("\r\n", "\n").replace("\r", "\n")
    t = re.sub(r"[ \t]+", " ", t).strip()

    # 문장 종결 표현 뒤 문단 분리
    t = re.sub(
        r"(다\.|했다\.|이다\.|고 밝혔다\.|라고 말했다\.)\s+(?=[가-힣A-Z0-9\"“‘])",
        r"\1\n\n",
        t
    )
    # 인용문 뒤 문단 분리
    t = re.sub(
        r"([\"”]\.?)\s+(?=[가-힣A-Z])",
        r"\1\n\n",
        t
    )
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t


def split_news_title_blocks(paras: list[str]) -> tuple[str, str, list[str]]:
    """
    뉴스 기사 문단 리스트를 제목/리드/본문으로 분리
    """
    if not paras:
        return "", "", []

    title_lines = []
    body_start = 0

    for i, p in enumerate(paras):
        s = (p or "").strip()
        if not s:
            continue

        # 출처 라인에서 제목 블록 종료
        if s.endswith("DB") or s.endswith("일보") or s.endswith("일보DB"):
            title_lines.append(s)
            body_start = i + 1
            break

        # 날짜/리드 시작 패턴(예: 6·3, 2026- 등)
        if s.startswith(("6·", "20", "19")):
            body_start = i
            break

        title_lines.append(s)
        if len(title_lines) >= 4:
            body_start = i + 1
            break

    title = " ".join(title_lines).strip()
    lead = paras[body_start].strip() if body_start < len(paras) else ""
    body = [x.strip() for x in paras[body_start + 1:] if (x or "").strip()] if body_start + 1 < len(paras) else []
    return title, lead, body

def save_to_file(content, default_name):
    if not content.strip():
        messagebox.showwarning("알림", "저장할 내용이 없습니다.")
        return

    filetypes = [
        ("텍스트 문서", "*.txt"),
        ("EPUB 전자책", "*.epub"),
        ("PDF 문서", "*.pdf"),
        ("Word 문서", "*.docx"),
        ("PowerPoint 문서", "*.pptx"),
        ("모든 파일", "*.*")
    ]
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        initialfile=default_name,
        filetypes=filetypes
    )
    if not file_path:
        return

    try:
        ext = os.path.splitext(file_path)[1].lower()
        paper_size = paper_size_var.get()

        dimensions = {
            "A4 (210x297mm)": (210, 297),
            "A5 (148x210mm)": (148, 210),
            "B5 (182x257mm)": (182, 257),
            "B6 (128x182mm)": (128, 182),
            "국판 (152x225mm)": (152, 225),
            "46판 (128x188mm)": (128, 188)
        }
        width, height = dimensions.get(paper_size, (210, 297))

        # 저장 상세 설정 스냅샷(저장 시점 적용 보장)
        d_font_fam = doc_font_family_var.get()
        d_font_sz = int(doc_font_size_var.get())
        d_line_sp = float(doc_line_spacing_var.get())
        d_para_sp = int(doc_para_spacing_var.get())
        d_marg_x = int(doc_margin_x_var.get())
        d_marg_y = int(doc_margin_y_var.get())
        d_page_num = bool(doc_page_numbers_var.get())
        d_indent_pt = int(doc_indent_pt_var.get())

        paragraphs_line = _split_lines_as_paragraphs(content)
        paragraphs_blank = _split_blanklines_as_paragraphs(content)

        if ext == ".docx":
            doc = docx.Document()

            style = doc.styles["Normal"]
            style.font.name = d_font_fam
            style.font.size = Pt(d_font_sz)
            style._element.rPr.rFonts.set(qn("w:eastAsia"), d_font_fam)

            for section in doc.sections:
                section.page_width = Mm(width)
                section.page_height = Mm(height)
                section.left_margin = Mm(d_marg_x)
                section.right_margin = Mm(d_marg_x)
                section.top_margin = Mm(d_marg_y)
                section.bottom_margin = Mm(d_marg_y)

                if d_page_num:
                    footer = section.footer
                    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run()
                    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
                    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
                    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
                    r._r.extend([fld1, instr, fld2])

            for text in paragraphs_line:
                para = doc.add_paragraph(text)
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                pf = para.paragraph_format
                pf.first_line_indent = Pt(d_indent_pt)
                pf.line_spacing = d_line_sp
                pf.space_after = Pt(d_para_sp)

            doc.save(file_path)

        elif ext == ".pdf":
            """
            ✅ PDF 저장 (Word 스타일 문단 첫 글자 들여쓰기)
            - 들여쓰기를 좌표로 하지 않고, 첫 줄에 '텍스트 공백'을 삽입
            - multi_cell()과 충돌 없음
            - 기사형 텍스트에서도 자연스럽게 동작
            """

            from fpdf import FPDF
            from fpdf.enums import XPos, YPos

            PT_TO_MM = 0.352777778

            class PDFWithFooter(FPDF):
                def __init__(self, *args, page_numbers=False, footer_font=("helvetica", 9), **kwargs):
                    super().__init__(*args, **kwargs)
                    self._page_numbers = page_numbers
                    self._footer_font_name = footer_font[0]
                    self._footer_font_size = footer_font[1]

                def footer(self):
                    if not self._page_numbers:
                        return
                    self.set_y(-12)
                    self.set_font(self._footer_font_name, size=self._footer_font_size)
                    self.cell(
                        0, 10, str(self.page_no()),
                        align="C",
                        new_x=XPos.LMARGIN,
                        new_y=YPos.NEXT
                    )

            # ── (1) PDF 객체 생성 ──────────────────────────
            font_path = _find_font_path_for_pdf(d_font_fam)
            footer_font_name = d_font_fam if font_path else "helvetica"
            footer_font_size = max(8, int(d_font_sz * 0.85))

            pdf = PDFWithFooter(
                format=(width, height),
                unit="mm",
                page_numbers=d_page_num,
                footer_font=(footer_font_name, footer_font_size)
            )

            pdf.set_margins(left=d_marg_x, top=d_marg_y, right=d_marg_x)
            pdf.set_auto_page_break(auto=True, margin=d_marg_y)
            pdf.add_page()

            # ── (2) 폰트 설정 ──────────────────────────────
            if font_path:
                pdf.add_font(d_font_fam, style="", fname=font_path)
                pdf.set_font(d_font_fam, size=d_font_sz)
            else:
                pdf.set_font("helvetica", size=d_font_sz)

            # ── (3) 줄 높이 계산 ───────────────────────────
            safe_line_spacing = max(1.0, float(d_line_sp))
            line_h = (pdf.font_size_pt * 1.25 * safe_line_spacing) * PT_TO_MM

            usable_w = width - (d_marg_x * 2)

            # ── (4) 문단 단위 출력 (소설책/문서형 PDF) ───────────
            for paragraph in paragraphs_line:
                pdf_write_paragraph_novel(
                    pdf,
                    paragraph,
                    page_w=width,
                    page_h=height,
                    margin_x=d_marg_x,
                    margin_y=d_marg_y,
                    base_indent_pt=d_indent_pt,
                    line_spacing=d_line_sp,
                    para_spacing_pt=d_para_sp
                )

            pdf.output(file_path)

        elif ext == ".epub":
            try:
                from ebooklib import epub
            except ImportError:
                messagebox.showerror("라이브러리 필요", "EPUB 저장을 위해 EbookLib가 필요합니다.\n\npip install EbookLib")
                return

            book = epub.EpubBook()
            book.set_title(default_name)
            book.set_language("ko")

            c1 = epub.EpubHtml(title="본문", file_name="chap_01.xhtml", lang="ko")
            html_content = (
                f'<html><head><meta charset="utf-8"></head>'
                f'<body style="margin: {d_marg_y}mm {d_marg_x}mm; font-family:\'{d_font_fam}\', sans-serif; word-break: break-all;">'
            )

            for text in paragraphs_line:
                html_content += (
                    f'<p style="text-indent: {d_indent_pt}pt; text-align: justify; line-height: {d_line_sp}; '
                    f'margin-bottom: {d_para_sp}pt; font-size: {d_font_sz}pt;">{html_mod.escape(text)}</p>'
                )
            html_content += "</body></html>"

            c1.content = html_content
            book.add_item(c1)
            book.toc = (c1,)
            book.add_item(epub.EpubNcx())
            book.add_item(epub.EpubNav())
            book.spine = ["nav", c1]
            epub.write_epub(file_path, book, {})

        elif ext == ".pptx":
            from pptx.enum.text import PP_ALIGN

            prs = Presentation()

            prs.slide_width, prs.slide_height = prs.slide_height, prs.slide_width
            blank_layout = prs.slide_layouts[6]

            EMU_PER_PT = 12700

            slide_w_emu = int(prs.slide_width)
            slide_h_emu = int(prs.slide_height)

            margin_x_emu = int(pptx.util.Mm(d_marg_x))
            margin_y_emu = int(pptx.util.Mm(d_marg_y))

            left = margin_x_emu
            top = margin_y_emu
            width_box_emu = max(1, slide_w_emu - margin_x_emu * 2)
            height_box_emu = max(1, slide_h_emu - margin_y_emu * 2)

            width_pt = float(width_box_emu) / EMU_PER_PT
            height_pt = float(height_box_emu) / EMU_PER_PT

            font_pt = float(d_font_sz)
            line_factor = max(1.1, float(d_line_sp)) * 1.25
            max_lines = max(8, int(height_pt / (font_pt * line_factor)))

            def estimate_chars_per_line():
                approx_char_w = font_pt * 0.90
                cpl = int(width_pt / max(1.0, approx_char_w))
                return max(20, min(70, cpl))

            chars_per_line = estimate_chars_per_line()

            def estimate_lines(text: str):
                if not text:
                    return 1
                eff = sum(0.55 if ord(c) < 128 else 1.0 for c in text)
                return max(1, int(math.ceil(eff / chars_per_line)))

            def create_new_slide():
                slide = prs.slides.add_slide(blank_layout)
                box = slide.shapes.add_textbox(left, top, width_box_emu, height_box_emu)
                tf = box.text_frame
                tf.clear()
                tf.word_wrap = True
                return tf

            tf = create_new_slide()
            current_lines = 0
            indent_spaces = _indent_spaces_from_pt(d_indent_pt)

            for text in paragraphs_line:
                need = estimate_lines(text) + 1

                if current_lines + need > max_lines and current_lines > 0:
                    tf = create_new_slide()
                    current_lines = 0

                p = tf.paragraphs[0] if not tf.paragraphs[0].text else tf.add_paragraph()
                p.text = (" " * indent_spaces) + text

                p.alignment = PP_ALIGN.JUSTIFY
                p.space_after = pptx.util.Pt(d_para_sp)
                p.line_spacing = d_line_sp
                p.font.size = pptx.util.Pt(d_font_sz)
                p.font.name = d_font_fam

                current_lines += need

            prs.save(file_path)

        else:
            indent_spaces = _indent_spaces_from_pt(d_indent_pt)
            with open(file_path, "w", encoding="utf-8") as f:
                for i, p in enumerate(paragraphs_line):
                    if i > 0:
                        f.write("\n\n")
                    f.write((" " * indent_spaces) + p)

        set_status("save", f"상태: 파일 저장 완료 ({os.path.basename(file_path)})", auto_revert_seconds=4)
        messagebox.showinfo("저장 완료", "파일이 성공적으로 저장되었습니다.")

    except Exception as e:
        set_status("error", "상태: 파일 저장 실패", auto_revert_seconds=6)
        messagebox.showerror("오류", f"저장 중 오류가 발생했습니다.\n{str(e)}")


def save_source_file():
    save_to_file(text_input.get("1.0", tk.END), "원문_추출내용")


def save_result_file():
    save_to_file(text_output.get("1.0", tk.END), "번역결과")


# ----------------- UI: 커스텀 메뉴 -----------------
class CustomMenuBar:
    def __init__(self, root_):
        self.frame = tk.Frame(root_, bg="#f0f0f0")
        self.frame.pack(side=tk.TOP, fill=tk.X)
        self.is_active = False
        self.active_menu = None
        self.menus = []
        root_.bind_all("<ButtonPress-1>", self.on_global_click, add="+")

    def activate_menu(self, menu):
        if self.active_menu and self.active_menu != menu:
            self.active_menu.hide()
        self.is_active = True
        self.active_menu = menu
        menu.show()

    def deactivate(self):
        if self.active_menu:
            self.active_menu.hide()
        self.is_active = False
        self.active_menu = None

    def on_global_click(self, event):
        if not self.is_active:
            return
        for m in self.menus:
            if event.widget == m.btn:
                return
        try:
            top = event.widget.winfo_toplevel()
            if self.active_menu and (top == self.active_menu.top or top in self.active_menu.submenus):
                return
        except Exception:
            pass
        self.deactivate()


class CustomDropdownMenu:
    def __init__(self, parent_bar, text, font):
        self.parent_bar = parent_bar
        self.parent_bar.menus.append(self)
        self.font = font

        self.btn = tk.Label(parent_bar.frame, text=text, font=font, bg="#f0f0f0", padx=12, pady=4, cursor="hand2")
        self.btn.pack(side=tk.LEFT)

        self.top = tk.Toplevel(parent_bar.frame)
        self.top.overrideredirect(True)
        self.top.withdraw()
        self.top.attributes("-topmost", True)

        self.frame = tk.Frame(self.top, bg="white", highlightbackground="#cccccc", highlightthickness=1)
        self.frame.pack(fill=tk.BOTH, expand=True)

        self.submenus = []
        self.items = []
        self.active_submenu = None

        def on_click(_e):
            if self.parent_bar.is_active and self.parent_bar.active_menu == self:
                self.parent_bar.deactivate()
            else:
                self.parent_bar.activate_menu(self)

        def on_enter(_e):
            if self.parent_bar.is_active and self.parent_bar.active_menu != self:
                self.parent_bar.activate_menu(self)
            elif not self.parent_bar.is_active:
                self.btn.config(bg="#cce8ff")

        def on_leave(_e):
            if not self.parent_bar.is_active or self.parent_bar.active_menu != self:
                self.btn.config(bg="#f0f0f0")

        self.btn.bind("<ButtonPress-1>", on_click)
        self.btn.bind("<Enter>", on_enter)
        self.btn.bind("<Leave>", on_leave)

    def show(self):
        self.btn.config(bg="#cce8ff")
        self.top.update_idletasks()
        x = self.btn.winfo_rootx()
        y = self.btn.winfo_rooty() + self.btn.winfo_height()
        self.top.geometry(f"+{x}+{y}")
        self.top.deiconify()
        self.top.lift()

    def hide(self):
        self.top.withdraw()
        for sub in self.submenus:
            sub.withdraw()
        self.active_submenu = None
        self.btn.config(bg="#f0f0f0")
        for item in self.items:
            item.config(bg="white")

    def add_command(self, label, command, is_special=False):
        color = "blue" if is_special else "black"
        item = tk.Label(self.frame, text=label, font=self.font, bg="white", fg=color, anchor="w", padx=15, pady=8)
        item.pack(fill=tk.X)
        self.items.append(item)

        def _enter(_e):
            for it in self.items:
                it.config(bg="white")
            item.config(bg="#cce8ff")
            if self.active_submenu:
                self.active_submenu.withdraw()
                self.active_submenu = None

        def _click(_e):
            self.parent_bar.deactivate()
            if command:
                self.btn.after(50, command)

        item.bind("<Enter>", _enter)
        item.bind("<ButtonRelease-1>", _click)

    def add_separator(self):
        tk.Frame(self.frame, bg="#eeeeee", height=1).pack(fill=tk.X, padx=5, pady=4)

    def add_cascade(self, label, values, var, suffix=""):
        item = tk.Label(self.frame, font=self.font, bg="white", anchor="w", padx=15, pady=6)
        item.pack(fill=tk.X)
        self.items.append(item)

        def update_parent_text(*_args):
            try:
                item.config(text=f"{label}: {var.get()}{suffix}  ▶")
            except Exception:
                pass

        var.trace_add("write", update_parent_text)
        update_parent_text()

        submenu = tk.Toplevel(self.top)
        submenu.overrideredirect(True)
        submenu.withdraw()
        submenu.attributes("-topmost", True)
        self.submenus.append(submenu)

        sub_frame = tk.Frame(submenu, bg="white", highlightbackground="#cccccc", highlightthickness=1)
        sub_frame.pack(fill=tk.BOTH, expand=True)

        sub_items = []

        def update_sub_text(lbl_widget, v_obj, val_str):
            lbl_widget.config(text=f"{'✔  ' if v_obj.get() == val_str else '   '}{val_str}{suffix}")

        for val in values:
            sub_item = tk.Label(sub_frame, font=self.font, bg="white", anchor="w", padx=15, pady=4)
            sub_item.pack(fill=tk.X)
            sub_items.append(sub_item)

            sub_item.bind("<ButtonRelease-1>", lambda e, v=var, vl=val: [v.set(vl), self.parent_bar.deactivate()])
            sub_item.bind("<Enter>", lambda e, w=sub_item, sibs=sub_items: [s.config(bg="white") for s in sibs] + [w.config(bg="#cce8ff")])

            var.trace_add("write", lambda *_a, l=sub_item, v=var, vl=val: update_sub_text(l, v, vl))
            update_sub_text(sub_item, var, val)

        def on_item_enter(_e):
            for it in self.items:
                it.config(bg="white")
            item.config(bg="#cce8ff")
            if self.active_submenu and self.active_submenu != submenu:
                self.active_submenu.withdraw()

            self.active_submenu = submenu
            self.top.update_idletasks()
            x = self.top.winfo_rootx() + self.top.winfo_width()
            y = item.winfo_rooty()
            submenu.geometry(f"+{x}+{y}")
            submenu.deiconify()
            submenu.lift()

        item.bind("<Enter>", on_item_enter)


def open_api_key_dialog():
    dialog = tk.Toplevel(root)
    dialog.title("API 키 설정 (공식 엔진 연동)")
    dialog.geometry("550x350")
    dialog.transient(root)
    dialog.grab_set()

    dialog.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - 225
    y = root.winfo_y() + (root.winfo_height() // 2) - 190
    dialog.geometry(f"+{x}+{y}")

    pad_frame = tk.Frame(dialog, padx=20, pady=10)
    pad_frame.pack(fill=tk.BOTH, expand=True)

    def create_api_entry(parent, label_text, default_val):
        tk.Label(parent, text=label_text, font=global_font, anchor="w").pack(fill=tk.X, pady=(10, 2))
        entry = tk.Entry(parent, width=45, font=global_font)
        entry.insert(0, default_val)
        entry.pack(fill=tk.X)
        return entry

    openai_entry = create_api_entry(pad_frame, "OpenAI API Key (gpt-4o-mini):", current_settings.get("openai_api_key", ""))
    deepl_entry = create_api_entry(pad_frame, "DeepL API Key (Free/Pro):", current_settings.get("deepl_api_key", ""))
    google_entry = create_api_entry(pad_frame, "Google Cloud Translation API Key:", current_settings.get("google_api_key", ""))
    ms_entry = create_api_entry(pad_frame, "MS Azure API Key (예: 키,지역코드 형식):", current_settings.get("ms_api_key", ""))

    def save_keys():
        current_settings["openai_api_key"] = openai_entry.get().strip()
        current_settings["deepl_api_key"] = deepl_entry.get().strip()
        current_settings["google_api_key"] = google_entry.get().strip()
        current_settings["ms_api_key"] = ms_entry.get().strip()
        save_current_config()
        set_status("ok", "상태: API 키가 적용되었습니다.", auto_revert_seconds=3)
        messagebox.showinfo("저장 완료", "API 키가 성공적으로 적용되었습니다.", parent=dialog)
        dialog.destroy()

    tk.Button(pad_frame, text="저장 및 닫기", font=global_font, command=save_keys,
              bg="#4CAF50", fg="white", width=15).pack(pady=(20, 0))


def open_save_options_dialog():
    dialog = tk.Toplevel(root)
    dialog.title("문서 저장 상세 설정")
    dialog.geometry("520x520")
    dialog.transient(root)
    dialog.grab_set()

    dialog.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - 260
    y = root.winfo_y() + (root.winfo_height() // 2) - 260
    dialog.geometry(f"+{x}+{y}")

    pad_frame = tk.Frame(dialog, padx=25, pady=20)
    pad_frame.pack(fill=tk.BOTH, expand=True)

    def create_spinbox_row(parent, label_text, current_val, from_val, to_val, inc_val, is_float=False, fmt="%.1f"):
        row = tk.Frame(parent)
        row.pack(fill=tk.X, pady=6)
        tk.Label(row, text=label_text, font=global_font, width=22, anchor="w").pack(side=tk.LEFT)
        if is_float:
            sb = ttk.Spinbox(row, from_=from_val, to=to_val, increment=inc_val, font=global_font, width=15, format=fmt)
        else:
            sb = ttk.Spinbox(row, from_=from_val, to=to_val, increment=inc_val, font=global_font, width=15)
        sb.set(current_val)
        sb.pack(side=tk.LEFT, padx=5)
        return sb

    def create_combobox_row(parent, label_text, current_val, values):
        row = tk.Frame(parent)
        row.pack(fill=tk.X, pady=6)
        tk.Label(row, text=label_text, font=global_font, width=22, anchor="w").pack(side=tk.LEFT)
        cb = ttk.Combobox(row, values=values, state="readonly", font=global_font, width=15)
        cb.set(current_val)
        cb.pack(side=tk.LEFT, padx=5)
        return cb

    def create_checkbox_row(parent, label_text, var_obj):
        row = tk.Frame(parent)
        row.pack(fill=tk.X, pady=6)
        tk.Label(row, text=label_text, font=global_font, width=22, anchor="w").pack(side=tk.LEFT)
        chk = ttk.Checkbutton(row, variable=var_obj, text="(Word, PDF 저장 시 적용)")
        chk.pack(side=tk.LEFT, padx=5)
        return chk

    ff_entry = create_combobox_row(pad_frame, "저장 글꼴:", doc_font_family_var.get(), ALLOWED_DOC_FONTS)
    fs_entry = create_spinbox_row(pad_frame, "저장 글자 크기 (pt):", doc_font_size_var.get(), 8, 72, 1)
    ls_entry = create_spinbox_row(pad_frame, "줄 간격 (배수):", doc_line_spacing_var.get(), 1.0, 3.0, 0.1, True, "%.1f")
    ps_entry = create_spinbox_row(pad_frame, "문단 뒤 간격 (pt):", doc_para_spacing_var.get(), 0, 60, 1)
    ind_entry = create_spinbox_row(pad_frame, "들여쓰기 (pt):", doc_indent_pt_var.get(), 0, 72, 1)
    mx_entry = create_spinbox_row(pad_frame, "좌우 여백 (mm):", doc_margin_x_var.get(), 5, 100, 1)
    my_entry = create_spinbox_row(pad_frame, "상하 여백 (mm):", doc_margin_y_var.get(), 5, 100, 1)
    create_checkbox_row(pad_frame, "하단 페이지 번호 표시:", doc_page_numbers_var)

    def save_doc_options():
        try:
            doc_font_family_var.set(ff_entry.get())
            doc_font_size_var.set(int(float(fs_entry.get())))
            doc_line_spacing_var.set(float(ls_entry.get()))
            doc_para_spacing_var.set(int(float(ps_entry.get())))
            doc_indent_pt_var.set(int(float(ind_entry.get())))
            doc_margin_x_var.set(int(float(mx_entry.get())))
            doc_margin_y_var.set(int(float(my_entry.get())))

            save_current_config()
            set_status("ok", "상태: 저장 세부 설정이 적용되었습니다.", auto_revert_seconds=3)
            messagebox.showinfo("저장 완료", "문서 저장 세부 설정이 적용되었습니다.", parent=dialog)
            dialog.destroy()
        except ValueError:
            messagebox.showerror("입력 오류", "숫자 형식으로 올바르게 입력해주세요.", parent=dialog)

    btn_row = tk.Frame(pad_frame)
    btn_row.pack(pady=(20, 0))

    tk.Button(
        btn_row,
        text="저장 및 닫기",
        font=global_font,
        command=save_doc_options,
        bg="#4CAF50",
        fg="white",
        width=15
    ).pack()


# ----------------- GUI 초기화 -----------------
root = TkinterDnD.Tk()
root.title("다중 엔진 문서 번역기 프로그램")
root.geometry("1100x800")

style = ttk.Style()
if "clam" in style.theme_names():
    style.theme_use("clam")
    style.configure("TSpinbox", arrowsize=24, padding=4)
    style.configure("TCombobox", arrowsize=20, padding=4)
    style.map("TCombobox",
              fieldbackground=[("readonly", "white")],
              selectbackground=[("readonly", "white"), ("focus", "white")],
              selectforeground=[("readonly", "black"), ("focus", "black")],
              foreground=[("readonly", "black")])
    style.map("TSpinbox",
              fieldbackground=[("readonly", "white")],
              selectbackground=[("readonly", "white"), ("focus", "white")],
              selectforeground=[("readonly", "black"), ("focus", "black")],
              foreground=[("readonly", "black")])

current_settings = load_settings()

font_size_var = tk.IntVar(value=current_settings["font_size"])
chunk_size_var = tk.IntVar(value=current_settings["chunk_size"])
engine_var = tk.StringVar(value=current_settings["engine"])
paper_size_var = tk.StringVar(value=current_settings["paper_size"])

doc_font_family_var = tk.StringVar(value=current_settings["doc_font_family"])
doc_font_size_var = tk.IntVar(value=current_settings["doc_font_size"])
doc_line_spacing_var = tk.DoubleVar(value=current_settings["doc_line_spacing"])
doc_para_spacing_var = tk.IntVar(value=current_settings["doc_para_spacing"])
doc_margin_x_var = tk.IntVar(value=current_settings["doc_margin_x"])
doc_margin_y_var = tk.IntVar(value=current_settings["doc_margin_y"])
doc_page_numbers_var = tk.BooleanVar(value=current_settings.get("doc_page_numbers", True))
doc_indent_pt_var = tk.IntVar(value=current_settings.get("doc_indent_pt", 14))

global_font = tkfont.Font(family="맑은 고딕", size=font_size_var.get())
bold_font = tkfont.Font(family="맑은 고딕", size=font_size_var.get(), weight="bold")
text_font = tkfont.Font(family=doc_font_family_var.get(), size=font_size_var.get())

tkfont.nametofont("TkDefaultFont").configure(family="맑은 고딕", size=font_size_var.get())
tkfont.nametofont("TkMenuFont").configure(family="맑은 고딕", size=font_size_var.get())
tkfont.nametofont("TkTextFont").configure(family="맑은 고딕", size=font_size_var.get())

root.option_add("*TCombobox*Listbox.font", global_font)
root.protocol("WM_DELETE_WINDOW", save_settings)


def update_font_size(*_args):
    new_size = font_size_var.get()
    global_font.configure(size=new_size)
    bold_font.configure(size=new_size)
    tkfont.nametofont("TkDefaultFont").configure(size=new_size)
    tkfont.nametofont("TkMenuFont").configure(size=new_size)
    tkfont.nametofont("TkTextFont").configure(size=new_size)
    text_font.configure(size=new_size)


def update_text_font_family(*_args):
    text_font.configure(family=doc_font_family_var.get())


font_size_var.trace_add("write", update_font_size)
doc_font_family_var.trace_add("write", update_text_font_family)

# ----------------- 상태바 먼저 만들고 BOTTOM 고정 -----------------
status_frame = tk.Frame(root, bd=1, relief=tk.SUNKEN)
status_frame.pack(side=tk.BOTTOM, fill=tk.X)

status_icon_var = tk.StringVar(value=STATUS_THEMES["idle"]["icon"])
status_var = tk.StringVar(value="상태: 대기 중")
metrics_var = tk.StringVar(value="")

status_icon_label = tk.Label(status_frame, textvariable=status_icon_var, font=global_font, padx=8, pady=5)
status_icon_label.pack(side=tk.LEFT)

status_label = tk.Label(status_frame, textvariable=status_var, font=global_font, anchor=tk.W, padx=4, pady=5)
status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

progress_var = tk.IntVar(value=0)
progress_label_var = tk.StringVar(value="0%")

progress_right = tk.Frame(status_frame)
progress_right.pack(side=tk.RIGHT, padx=10)

progressbar = ttk.Progressbar(progress_right, orient="horizontal", length=160, mode="determinate",
                              maximum=100, variable=progress_var)
progressbar.pack(side=tk.LEFT, padx=(0, 8), pady=4)

progress_label = tk.Label(progress_right, textvariable=progress_label_var, font=global_font)
progress_label.pack(side=tk.LEFT, padx=(0, 10))

metrics_label = tk.Label(progress_right, textvariable=metrics_var, font=global_font, anchor=tk.E)
metrics_label.pack(side=tk.LEFT)

_apply_status_theme("idle")
_set_progress_percent(0)

# ----------------- content 영역(상태바 위) -----------------
content_container = tk.Frame(root)
content_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# ----------------- 커스텀 메뉴 바 -----------------
menubar = CustomMenuBar(content_container)

file_menu = CustomDropdownMenu(menubar, "파일(F)", global_font)
file_menu.add_command("문서 파일 불러오기...", load_from_file)
file_menu.add_command("웹사이트 텍스트로 불러오기...", load_from_url)
file_menu.add_separator()
file_menu.add_command("전체 초기화", on_reset)
file_menu.add_command("종료", save_settings)

save_menu = CustomDropdownMenu(menubar, "저장(S)", global_font)
save_menu.add_command("좌측 원문 저장", save_source_file)
save_menu.add_command("우측 번역본 저장", save_result_file)

setting_menu = CustomDropdownMenu(menubar, "설정(O)", global_font)
setting_menu.add_cascade("번역 엔진", ALLOWED_ENGINES, engine_var, "")
setting_menu.add_cascade("글자 크기 (화면용)", ALLOWED_FONT_SIZES, font_size_var, "pt")
setting_menu.add_cascade("번역 단위", ALLOWED_CHUNK_SIZES, chunk_size_var, "자")
setting_menu.add_cascade("문서 저장 기본 용지", ALLOWED_PAPER_SIZES, paper_size_var, "")
# (요청 반영) PDF 저장 프리셋 / 한국 소설책 프리셋 메뉴 항목 삭제
setting_menu.add_separator()
setting_menu.add_command("[ 문서 저장 상세 설정 ]", open_save_options_dialog, is_special=False)
setting_menu.add_command("[ API 키 설정 (공식 엔진) ]", open_api_key_dialog, is_special=True)

# ----------------- 리본 메뉴 -----------------
ribbon_frame = tk.Frame(content_container, bg="#f5f6f7", bd=1, relief=tk.RAISED, pady=6, padx=5)
ribbon_frame.pack(side=tk.TOP, fill=tk.X)

def create_ribbon_btn(parent, text, command, highlight=False):
    bg_color = "#4CAF50" if highlight else "#ffffff"
    fg_color = "white" if highlight else "black"
    btn = tk.Button(parent, text=text, font=global_font, bg=bg_color, fg=fg_color,
                    relief=tk.RAISED if highlight else tk.GROOVE, command=command, padx=8, pady=2)
    btn.pack(side=tk.LEFT, padx=3, pady=2, ipady=3)
    return btn

tk.Frame(ribbon_frame, width=2, bg="#cccccc").pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=3)
create_ribbon_btn(ribbon_frame, "📁 파일 열기", load_from_file)
create_ribbon_btn(ribbon_frame, "🌐 웹 열기", load_from_url)
tk.Frame(ribbon_frame, width=2, bg="#cccccc").pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=3)
create_ribbon_btn(ribbon_frame, "💾 원문 저장", save_source_file)
create_ribbon_btn(ribbon_frame, "💾 결과 저장", save_result_file)
tk.Frame(ribbon_frame, width=2, bg="#cccccc").pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=3)
create_ribbon_btn(ribbon_frame, "🔄 화면 초기화", on_reset)

btn_translate = create_ribbon_btn(ribbon_frame, "▶ 번역 실행", on_translate, highlight=True)
btn_cancel = create_ribbon_btn(ribbon_frame, "■ 중단", on_cancel_translate, highlight=False)
btn_cancel.config(state=tk.DISABLED)

tk.Frame(content_container, height=1, bg="#d3d3d3").pack(fill=tk.X)

# ----------------- 상단 옵션 바 -----------------
top_frame = tk.Frame(content_container, pady=10, padx=10)
top_frame.pack(fill=tk.X)

tk.Label(top_frame, text="번역 엔진:", font=global_font).pack(side=tk.LEFT, padx=(0, 5))
engine_combobox = ttk.Combobox(top_frame, values=ALLOWED_ENGINES, state="readonly", width=16, font=global_font)
engine_combobox.set(engine_var.get())
engine_combobox.pack(side=tk.LEFT, padx=(0, 15))
engine_combobox.bind("<<ComboboxSelected>>", lambda e: engine_var.set(engine_combobox.get()))
engine_var.trace_add("write", lambda *_a: engine_combobox.set(engine_var.get()))

tk.Label(top_frame, text="도착 언어:", font=global_font).pack(side=tk.LEFT, padx=(0, 5))
lang_combobox = ttk.Combobox(top_frame, values=list(LANGUAGES.keys()), state="readonly", width=12, font=global_font)
lang_combobox.set(current_settings["target_lang"])
lang_combobox.pack(side=tk.LEFT, padx=(0, 15))

tk.Label(top_frame, text="단위:", font=global_font).pack(side=tk.LEFT, padx=(0, 5))
chunk_combobox = ttk.Combobox(top_frame, values=[f"{c}자" for c in ALLOWED_CHUNK_SIZES],
                              state="readonly", width=6, font=global_font)
chunk_combobox.set(f"{current_settings['chunk_size']}자")
chunk_combobox.pack(side=tk.LEFT, padx=(0, 15))
chunk_combobox.bind("<<ComboboxSelected>>", lambda e: chunk_size_var.set(int(chunk_combobox.get().replace("자", ""))))
chunk_size_var.trace_add("write", lambda *_a: chunk_combobox.set(f"{chunk_size_var.get()}자"))

tk.Label(top_frame, text="출판 용지:", font=global_font).pack(side=tk.LEFT, padx=(0, 5))
paper_combobox = ttk.Combobox(top_frame, values=ALLOWED_PAPER_SIZES, state="readonly", width=18, font=global_font)
paper_combobox.set(paper_size_var.get())
paper_combobox.pack(side=tk.LEFT, padx=(0, 5))
paper_combobox.bind("<<ComboboxSelected>>", lambda e: paper_size_var.set(paper_combobox.get()))
paper_size_var.trace_add("write", lambda *_a: paper_combobox.set(paper_size_var.get()))

# ----------------- 메인 텍스트 영역 -----------------
main_frame = tk.Frame(content_container)
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
main_frame.columnconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)
main_frame.rowconfigure(0, weight=1)

novel_view_config = {
    "wrap": tk.WORD,
    "spacing1": 8,
    "spacing2": 5,
    "spacing3": 12,
    "padx": 25,
    "pady": 20,
    "bg": "#FDFCF7"
}

is_syncing = False

def scroll_both(*args):
    text_input.yview(*args)
    text_output.yview(*args)

def sync_scroll_input(*args):
    global is_syncing
    input_scroll.set(*args)
    if not is_syncing:
        is_syncing = True
        text_output.yview_moveto(args[0])
        is_syncing = False

def sync_scroll_output(*args):
    global is_syncing
    output_scroll.set(*args)
    if not is_syncing:
        is_syncing = True
        text_input.yview_moveto(args[0])
        is_syncing = False

input_frame = tk.Frame(main_frame)
input_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

tk.Label(input_frame, text="[ 원문 ]", font=global_font, anchor="w").pack(fill=tk.X)

input_scroll = tk.Scrollbar(input_frame)
input_scroll.pack(side=tk.RIGHT, fill=tk.Y)

text_input = tk.Text(input_frame, font=text_font, **novel_view_config)
text_input.pack(fill=tk.BOTH, expand=True)

text_input.config(yscrollcommand=sync_scroll_input)
input_scroll.config(command=scroll_both)

text_input.drop_target_register(DND_FILES)
text_input.dnd_bind("<<Drop>>", handle_drop)

output_frame = tk.Frame(main_frame)
output_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

tk.Label(output_frame, text="[ 번역 결과 ]", font=global_font, anchor="w").pack(fill=tk.X)

output_scroll = tk.Scrollbar(output_frame)
output_scroll.pack(side=tk.RIGHT, fill=tk.Y)

text_output = tk.Text(output_frame, font=text_font, **novel_view_config)
text_output.pack(fill=tk.BOTH, expand=True)

text_output.config(yscrollcommand=sync_scroll_output)
output_scroll.config(command=scroll_both)

sync_timer = None

def sync_highlight(source, target):
    try:
        start = source.index(tk.SEL_FIRST)
        end = source.index(tk.SEL_LAST)
        target.tag_remove("sync_hl", "1.0", tk.END)

        s_line = int(start.split(".")[0])
        e_line = int(end.split(".")[0])
        max_line = min(e_line, s_line + 200)

        for line in range(s_line, max_line + 1):
            src_text = source.get(f"{line}.0", f"{line}.end")
            tgt_text = target.get(f"{line}.0", f"{line}.end")
            if not src_text or not tgt_text:
                continue

            ratio = len(tgt_text) / max(1, len(src_text))
            target.tag_add(
                "sync_hl",
                f"{line}.{int(0 * ratio)}",
                f"{line}.{int(len(src_text) * ratio)}"
            )

        target.tag_config("sync_hl", background="#DDEBF7")
    except tk.TclError:
        target.tag_remove("sync_hl", "1.0", tk.END)

def on_select_input(_e):
    global sync_timer
    if sync_timer:
        root.after_cancel(sync_timer)
    sync_timer = root.after(120, lambda: sync_highlight(text_input, text_output))

def on_select_output(_e):
    global sync_timer
    if sync_timer:
        root.after_cancel(sync_timer)
    sync_timer = root.after(120, lambda: sync_highlight(text_output, text_input))

text_input.bind("<<Selection>>", on_select_input)
text_output.bind("<<Selection>>", on_select_output)

root.after(60, _poll_translate_queue)

set_status("idle", "상태: 대기 중")
root.mainloop()