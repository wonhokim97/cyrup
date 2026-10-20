import sys
import os
import json
import time
import threading
import ctypes
import pygetwindow as gw
import tkinter as tk
from tkinter import ttk, messagebox, font, simpledialog, filedialog
import pyautogui
from pynput import mouse, keyboard
import win32gui

# =========================================================================
# 1. 데이터 관리 클래스 (DataManager)
# =========================================================================
class DataManager:
    def __init__(self):
        self.config_path = os.path.join(os.getcwd(), "smart_macro_config.json")
        self.macro_db_path = os.path.join(os.getcwd(), "smart_macros_db.json")

    def load_config(self):
        default_config = {
            "target_window": "",
            "loop_count": "1",
            "loop_delay": "1.0",
            "font_size": 10,
            "geometry": "1050x650",
            "auto_correct": True,
            "dynamic_size_wait": True,
            "uia_search": True,
            "popup_link": True,
            "edit_win_geom": "380x400",
            "tree_column_widths": [45, 60, 75, 75, 95, 230]
        }
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    default_config.update(loaded)
            except (IOError, json.JSONDecodeError) as e:
                print(f"설정 로드 오류: {e}")
        return default_config

    def save_config(self, config_data):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
        except IOError as e:
            print(f"설정 저장 실패: {e}")

    def load_macro_db(self):
        if os.path.exists(self.macro_db_path):
            try:
                with open(self.macro_db_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (IOError, json.JSONDecodeError):
                pass
        return {}

    def save_macro_db(self, db_data):
        try:
            with open(self.macro_db_path, "w", encoding="utf-8") as f:
                json.dump(db_data, f, ensure_ascii=False, indent=4)
        except IOError as e:
            messagebox.showerror("에러", f"매크로 데이터 저장 실패: {e}")

# =========================================================================
# 2. 매크로 기록 및 실행 엔진 (MacroEngine)
# =========================================================================
class MacroEngine:
    def __init__(self, controller):
        self.controller = controller
        self.is_recording = False
        self.is_running = False
        self.macro_list = []
        self.last_click_time = None
        self.mouse_listener = None
        self.keyboard_listener = None
        self.first_click_window_title = None
        self.first_click_record_saved = False
        self.last_recorded_click = None
        self.pending_mouse_press = None
        self.start_keyboard_listener()

    def normalize_target_title(self, title):
        if title is None:
            return ""
        cleaned = str(title).strip()
        if not cleaned:
            return ""
        app_title = self.controller.root.title().strip()
        if cleaned == app_title:
            return ""
        return cleaned

    def start_keyboard_listener(self):
        def on_press(key):
            try:
                if key == keyboard.Key.f2:
                    self.controller.root.after(0, self.controller.toggle_recording)
                elif key == keyboard.Key.f3:
                    self.controller.root.after(0, self.controller.toggle_running)
                elif key == keyboard.Key.enter and self.is_recording:
                    self.on_key_press(key)
            except Exception:
                pass
        self.keyboard_listener = keyboard.Listener(on_press=on_press)
        self.keyboard_listener.start()

    def on_key_press(self, key):
        if not self.is_recording:
            return

        try:
            active_win = gw.getActiveWindow()
            app_window_title = self.controller.root.title().strip()
            if active_win and active_win.title.strip() == app_window_title:
                return

            current_time = time.time()
            delay = current_time - self.last_click_time if self.last_click_time else 0.0
            self.last_click_time = current_time
            self.last_recorded_click = None

            recorded_item = ("key_press", "enter", "", round(delay, 2), "")
            self.macro_list.append(recorded_item)
            self.controller.safe_log(f"엔터키 기록됨, 딜레이: {round(delay, 2)}초")
            self.controller.root.after(0, lambda: self.controller.ui.update_treeview(self.macro_list))
        except Exception as e:
            print(f"엔터키 기록 중 예외: {e}")

    def toggle_recording(self):
        if not self.is_recording:
            if self.macro_list:
                result = messagebox.askyesno(
                    "기록 확인",
                    "매크로 동작 편집에 이미 마우스 클릭 정보가 있습니다.\n"
                    "새로 기록하기 전에 [초기화]를 실행하는 것을 권장합니다.\n\n"
                    "기존 기록을 모두 지우고 새로 기록을 시작하시겠습니까?",
                    icon='warning'
                )
                if not result:
                    return False

            try:
                active_window = gw.getActiveWindow()
                if not active_window or not active_window.title.strip():
                    messagebox.showwarning("경고", "현재 활성화된 유효한 창을 찾을 수 없습니다.")
                    return False
            except Exception as e:
                messagebox.showerror("에러", f"창 확인 중 오류 발생: {e}")
                return False

            self.is_recording = True
            self.macro_list.clear()
            self.first_click_window_title = None
            self.first_click_record_saved = False
            self.last_recorded_click = None
            self.pending_mouse_press = None
            
            self.controller.config["target_window"] = ""
            self.controller.root.after(0, lambda: self.controller.ui.set_target_window(""))
            self.controller.root.after(0, lambda: self.controller.ui.update_treeview(self.macro_list))

            self.last_click_time = time.time()
            self.controller.safe_log("매크로 기록 시작됨. 팝업창을 포함하여 화면을 클릭하세요.")

            self.mouse_listener = mouse.Listener(on_click=self.on_mouse_click)
            self.mouse_listener.start()
            return True
        else:
            self.stop_recording()
            return False

    def stop_recording(self):
        if self.is_recording:
            self.is_recording = False
            self.last_recorded_click = None
            self.pending_mouse_press = None
            if self.mouse_listener:
                self.mouse_listener.stop()
            self.controller.root.after(0, lambda: self.controller.ui.update_treeview(self.macro_list))
            self.controller.safe_log("매크로 기록이 중지되었습니다.")

    def on_mouse_click(self, x, y, button, pressed):
        if not self.is_recording:
            return

        if pressed:
            self.pending_mouse_press = (button, x, y)
            return

        if not self.pending_mouse_press:
            return

        pressed_button, pressed_x, pressed_y = self.pending_mouse_press
        self.pending_mouse_press = None
        if (
            button != pressed_button
            or abs(x - pressed_x) > 5
            or abs(y - pressed_y) > 5
        ):
            return

        if self.is_recording:
            try:
                app_window_title = self.controller.root.title().strip()
                active_win = gw.getActiveWindow()
                if active_win and active_win.title.strip() == app_window_title:
                    # 프로그램 본체의 UI 영역 클릭은 매크로 대상이 아니므로 기록에서 제외
                    # 이렇게 해야 팝업 연동 OFF 상태에서도 본체 창으로만 좌표가 잡히는 문제를 방지할 수 있음
                    return

                current_time = time.time()
                double_click_time = ctypes.windll.user32.GetDoubleClickTime() / 1000.0
                if (
                    self.last_recorded_click
                    and current_time - self.last_recorded_click[0] <= double_click_time
                    and self.last_recorded_click[1] == button
                    and abs(x - self.last_recorded_click[2]) <= 5
                    and abs(y - self.last_recorded_click[3]) <= 5
                ):
                    previous_index = self.last_recorded_click[4]
                    previous_item = self.macro_list[previous_index]
                    self.macro_list[previous_index] = ("double_click",) + tuple(previous_item[1:])
                    self.last_recorded_click = None
                    self.last_click_time = current_time
                    self.controller.safe_log(f"더블 클릭으로 변경됨: ({x}, {y})")
                    self.controller.root.after(0, lambda: self.controller.ui.update_treeview(self.macro_list))
                    return

                self.last_recorded_click = None

                win_title = ""
                win_left = 0
                win_top = 0

                # 1단계: 현재 활성 창에서 제목 감지
                if active_win and active_win.title.strip():
                    detected_title = active_win.title.strip()
                    if detected_title != app_window_title:
                        win_title = detected_title
                        win_left = active_win.left
                        win_top = active_win.top

                # 2단계: 활성 창이 없거나 앱 자신이면, 저장된 기본 창 제목 사용
                if not win_title:
                    stored_title = self.controller.normalize_target_window_title(self.controller.ui.get_target_window())
                    if stored_title:
                        win_title = stored_title

                # 3단계: 여전히 제목이 없으면 마우스 좌표로부터 창 탐지 시도
                if not win_title:
                    try:
                        # pynput 대신 pyautogui 사용하여 마우스 위치 재확인
                        import win32gui
                        hwnd = win32gui.WindowFromPoint((x, y))
                        if hwnd:
                            window_title = win32gui.GetWindowText(hwnd)
                            if window_title and window_title.strip() and window_title != app_window_title:
                                win_title = window_title.strip()
                                try:
                                    rect = win32gui.GetWindowRect(hwnd)
                                    win_left, win_top = rect[0], rect[1]
                                except Exception:
                                    pass
                    except Exception:
                        pass

                # 4단계: 첫 클릭 시 기본 창 제목 자동 기록
                if self.first_click_window_title is None and win_title and win_title != app_window_title:
                    self.first_click_window_title = win_title
                    self.controller.ui.set_target_window(win_title)
                    self.controller.config["target_window"] = win_title
                    self.controller.safe_log(f"기본 창 제목 자동 기록: '{win_title}'")

                delay = current_time - self.last_click_time if self.last_click_time else 0.0
                self.last_click_time = current_time

                rel_x = x - win_left
                rel_y = y - win_top
                recorded_item = ("click", rel_x, rel_y, round(delay, 2), win_title, x, y)

                self.macro_list.append(recorded_item)
                self.last_recorded_click = (current_time, button, x, y, len(self.macro_list) - 1)
                if self.first_click_window_title is None and win_title:
                    self.first_click_window_title = win_title
                self.controller.safe_log(f"기록됨 [{win_title or '미확인'}] 좌표: ({rel_x}, {rel_y}), 절대좌표: ({x}, {y}), 딜레이: {round(delay, 2)}초")
                self.controller.root.after(0, lambda: self.controller.ui.update_treeview(self.macro_list))
            except Exception as e:
                print(f"클릭 기록 중 예외: {e}")

    def toggle_running(self):
        if not self.is_running:
            if not self.macro_list:
                messagebox.showwarning("경고", "실행할 매크로 기록이 없습니다.")
                return False
            
            self.is_running = True
            threading.Thread(target=self.run_macro_thread, daemon=True).start()
            return True
        else:
            self.is_running = False
            self.controller.safe_log("매크로 실행 중단됨.")
            return False

    def perform_click(self, x, y, action_type):
        if action_type == "double_click":
            pyautogui.doubleClick(x, y, interval=0.1)
        else:
            pyautogui.click(x, y)

    def run_macro_thread(self):
        fallback_target = self.controller.ui.get_target_window()
        is_auto_correct = self.controller.ui.get_auto_correct()
        is_dynamic_wait = self.controller.ui.get_dynamic_size_wait()
        is_uia_search = self.controller.ui.get_uia_search()
        is_popup_link = self.controller.ui.get_popup_link()
        app_window_title = self.controller.root.title().strip()
        
        try:
            max_loops = int(self.controller.ui.get_loop_count())
        except ValueError:
            max_loops = 1

        try:
            loop_delay = float(self.controller.ui.get_loop_delay())
        except ValueError:
            loop_delay = 1.0

        self.controller.safe_log("매크로 실행을 시작합니다... (F3으로 중단 가능)")

        current_loop = 0
        while self.is_running:
            if max_loops > 0 and current_loop >= max_loops:
                break
            
            current_loop += 1
            loop_str = f"{current_loop}/{max_loops}" if max_loops > 0 else f"{current_loop}/무한"
            self.controller.safe_log(f"매크로 반복 실행 중 ({loop_str})...")

            try:
                for idx, action in enumerate(self.macro_list):
                    if not self.is_running:
                        break
                    
                    if len(action) >= 7:
                        act_type, rel_x, rel_y, delay, action_win_title, abs_x, abs_y = action[:7]
                    elif len(action) >= 5:
                        act_type, rel_x, rel_y, delay, action_win_title = action[:5]
                        abs_x = None
                        abs_y = None
                    else:
                        act_type, rel_x, rel_y, delay = action
                        action_win_title = ""
                        abs_x = None
                        abs_y = None

                    if delay > 0:
                        elapsed = 0
                        while elapsed < delay:
                            if not self.is_running:
                                break
                            time.sleep(0.05)
                            elapsed += 0.05

                    if not self.is_running:
                        break

                    if act_type == "key_press":
                        pyautogui.press(rel_x or "enter")
                        self.controller.safe_log(f"실행 [{idx+1}] 엔터키 입력")
                        continue
                    
                    # 대상 창 결정
                    # 1) 팝업 연동 ON: 기록된 팝업 제목 우선 사용
                    # 2) 기본 창 제목은 항상 보존하며, 팝업 연동 OFF라도 실제 실행 대상의 기준으로 사용
                    # 3) 타이틀이 없을 때만 현재 활성 창을 기준으로 실행
                    recorded_popup_title = action_win_title.strip() if action_win_title else ""
                    recorded_basic_title = fallback_target.strip() if fallback_target else ""

                    if recorded_popup_title and is_popup_link:
                        target_title_to_find = recorded_popup_title
                        search_mode = "popup"
                    elif recorded_basic_title:
                        target_title_to_find = recorded_basic_title
                        search_mode = "basic"
                    elif recorded_popup_title:
                        target_title_to_find = recorded_popup_title
                        search_mode = "popup"
                    else:
                        target_title_to_find = None
                        search_mode = "active"

                    if not is_popup_link:
                        self.controller.safe_log(
                            "팝업 연동 OFF: 제목 대신 좌표와 시간 정보를 기준으로 실행합니다."
                        )
                        self.controller.safe_log(
                            f"실행 좌표 기준 모드: ({rel_x}, {rel_y}), 대기시간 {delay}초"
                        )
                        if recorded_basic_title:
                            self.controller.safe_log(f"기본 창 제목 기록 유지: '{recorded_basic_title}' (실행에는 사용하지 않음)")

                    windows = []
                    timeout_seconds = 4.0
                    wait_deadline = time.monotonic() + timeout_seconds
                    last_wait_log = 0.0

                    if not is_popup_link:
                        try:
                            self.perform_click(rel_x, rel_y, act_type)
                            self.controller.safe_log(f"실행 [{idx+1}] 기록된 화면 좌표 클릭: ({rel_x}, {rel_y})")
                        except Exception as e:
                            self.controller.safe_log(f"실행 [{idx+1}] 기록된 좌표 클릭 오류: {e}")
                        continue

                    # 타겟 창 검색
                    if search_mode == "popup":
                        while self.is_running:
                            try:
                                if target_title_to_find and target_title_to_find.strip():
                                    windows = gw.getWindowsWithTitle(target_title_to_find)
                                if windows:
                                    break
                            except Exception:
                                windows = []

                            now = time.monotonic()
                            if now >= wait_deadline:
                                if fallback_target and fallback_target.strip():
                                    if is_popup_link:
                                        self.controller.safe_log(
                                            f"팝업 창 '{target_title_to_find}' 미검출, 기본 창 '{fallback_target}'로 대체합니다."
                                        )
                                    target_title_to_find = fallback_target
                                    wait_deadline = time.monotonic() + 2.0
                                else:
                                    if is_popup_link:
                                        self.controller.safe_log(
                                            f"팝업 창 '{target_title_to_find}' 미검출, 현재 활성 창에서 실행합니다."
                                        )
                                    break

                            if now - last_wait_log >= 1.0:
                                if is_popup_link:
                                    self.controller.safe_log(f"대기 중: '{target_title_to_find}' 창을 찾는 중...")
                                last_wait_log = now
                            time.sleep(0.2)
                    elif search_mode == "basic":
                        try:
                            if target_title_to_find and target_title_to_find.strip():
                                windows = gw.getWindowsWithTitle(target_title_to_find)
                        except Exception:
                            windows = []
                    else:
                        # 팝업 연동 OFF 또는 기본 창 정보 없음: 현재 활성 창 기반으로 즉시 실행
                        try:
                            active_win = gw.getActiveWindow()
                            if active_win and active_win.title.strip() and active_win.title.strip() != app_window_title:
                                windows = [active_win]
                                if is_popup_link:
                                    self.controller.safe_log("현재 활성 창을 기준으로 실행합니다.")
                        except Exception:
                            windows = []

                    # 창을 찾지 못했을 때 현재 활성 창 사용
                    if not windows:
                        try:
                            active_win = gw.getActiveWindow()
                            if active_win and active_win.title.strip() and active_win.title.strip() != app_window_title:
                                windows = [active_win]
                                if is_popup_link and target_title_to_find:
                                    self.controller.safe_log(
                                        "대상 창을 찾지 못해 현재 활성 창에서 실행합니다."
                                    )
                        except Exception:
                            windows = []

                    if not self.is_running:
                        break

                    if not windows:
                        if abs_x is not None and abs_y is not None:
                            self.perform_click(abs_x, abs_y, act_type)
                            self.controller.safe_log(f"실행 [{idx+1}] 대상 창을 찾지 못해 절대 좌표 클릭: ({abs_x}, {abs_y})")
                        else:
                            self.controller.safe_log(f"실행 [{idx+1}] 대상 창을 찾지 못해 동작 건너뜀")
                        continue
                    
                    try:
                        windows = sorted(windows, key=lambda w: (0 if w.title == target_title_to_find else 1, len(w.title)))
                        current_window = windows[0]
                    except Exception:
                        continue
                    
                    try:
                        if hasattr(current_window, 'isMinimized') and current_window.isMinimized:
                            current_window.restore()
                            time.sleep(0.1)
                        
                        if not current_window.isActive:
                            current_window.activate()
                    except Exception as e:
                        print(f"창 활성화/복원 중 예외: {e}")

                    if is_dynamic_wait:
                        try:
                            w_prev, h_prev = current_window.width, current_window.height
                            for _ in range(10):
                                time.sleep(0.05)
                                w_curr, h_curr = current_window.width, current_window.height
                                if w_curr == w_prev and h_curr == h_prev:
                                    break
                                w_prev, h_prev = w_curr, h_curr
                        except Exception:
                            pass

                    target_abs_x = None
                    target_abs_y = None

                    if is_uia_search and target_title_to_find and target_title_to_find.strip():
                        try:
                            from pywinauto import Application
                            app = Application(backend="uia").connect(title_re=f".*{target_title_to_find}.*")
                            win_uia = app.window(title_re=f".*{target_title_to_find}.*")
                            rect = win_uia.rectangle()
                            if rect:
                                target_abs_x = rect.left + rel_x
                                target_abs_y = rect.top + rel_y
                        except Exception:
                            pass

                    if target_abs_x is None or target_abs_y is None:
                        if is_auto_correct:
                            curr_x, curr_y = current_window.left, current_window.top
                            target_abs_x = curr_x + rel_x
                            target_abs_y = curr_y + rel_y
                        else:
                            target_abs_x = current_window.left + rel_x
                            target_abs_y = current_window.top + rel_y

                    self.perform_click(target_abs_x, target_abs_y, act_type)
                    if is_popup_link:
                        self.controller.safe_log(f"실행 [{idx+1}] [{target_title_to_find}] 클릭: ({target_abs_x}, {target_abs_y})")
                    else:
                        self.controller.safe_log(f"실행 [{idx+1}] 좌표 클릭: ({target_abs_x}, {target_abs_y})")
                
                if not self.is_running:
                    break

                if max_loops == 0 or current_loop < max_loops:
                    self.controller.safe_log(f"회차 간 대기 중 ({loop_delay}초)...")
                    elapsed = 0
                    while elapsed < loop_delay:
                        if not self.is_running:
                            break
                        time.sleep(0.05)
                        elapsed += 0.05

            except Exception as e:
                self.controller.safe_log(f"실행 중 예외 발생: {e}")
                break

        self.is_running = False
        self.controller.root.after(0, self.controller.ui.set_run_button_stopped)
        self.controller.safe_log("매크로 실행이 종료되었습니다.")


# =========================================================================
# 3. 사용자 인터페이스 (MacroAppUI)
# =========================================================================
class MacroAppUI:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        
        self.current_font_size = self.controller.config.get("font_size", 10)
        self.app_font = font.Font(family="Malgun Gothic", size=self.current_font_size)
        self.root.option_add("*Font", self.app_font)

        self.auto_correct_var = tk.BooleanVar(value=self.controller.config.get("auto_correct", True))
        self.dynamic_size_wait_var = tk.BooleanVar(value=self.controller.config.get("dynamic_size_wait", True))
        self.uia_search_var = tk.BooleanVar(value=self.controller.config.get("uia_search", True))
        self.popup_link_var = tk.BooleanVar(value=self.controller.config.get("popup_link", True))

        self.style = ttk.Style()
        self.update_styles()

        self.menubar = None
        self.create_menu()
        self.create_widgets()

    def update_styles(self):
        # 폰트 크기 10을 기준으로 최적 행 높이(24)를 설정하고, 글자 크기에 비례하여 자동 조정
        dynamic_rowheight = int(24 * (self.current_font_size / 10.0))
        self.style.configure("Treeview", font=("Malgun Gothic", self.current_font_size), rowheight=dynamic_rowheight)
        self.style.configure("Treeview.Heading", font=("Malgun Gothic", self.current_font_size, "bold"))

    def create_menu(self):
        if self.menubar:
            self.menubar.destroy()

        self.menubar = tk.Menu(self.root, font=self.app_font)

        file_menu = tk.Menu(self.menubar, tearoff=0, font=self.app_font)
        file_menu.add_command(label="매크로 저장 (새 이름)", command=self.controller.save_macro)
        file_menu.add_command(label="매크로 DB 가져오기", command=self.controller.load_macro_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.controller.on_close)
        self.menubar.add_cascade(label="파일", menu=file_menu)

        edit_menu = tk.Menu(self.menubar, tearoff=0, font=self.app_font)
        edit_menu.add_command(label="선택 삭제", command=self.controller.delete_selected_macro)
        edit_menu.add_separator()
        edit_menu.add_command(label="초기화", command=self.controller.clear_macro)
        self.menubar.add_cascade(label="편집", menu=edit_menu)

        view_menu = tk.Menu(self.menubar, tearoff=0, font=self.app_font)
        view_menu.add_command(label="글자 크게 (+)", command=lambda: self.controller.change_font_size(2))
        view_menu.add_command(label="글자 작게 (-)", command=lambda: self.controller.change_font_size(-2))
        self.menubar.add_cascade(label="보기", menu=view_menu)

        help_menu = tk.Menu(self.menubar, tearoff=0, font=self.app_font)
        help_menu.add_command(label="프로그램 도움말", command=self.controller.show_help_dialog)
        self.menubar.add_cascade(label="도움말", menu=help_menu)

        self.root.config(menu=self.menubar)

    def create_widgets(self):
        toolbar_frame = tk.Frame(self.root, bd=1, relief=tk.RAISED, bg="#f7f7f7")
        toolbar_frame.pack(side="top", fill="x", padx=5, pady=5)

        tk.Label(toolbar_frame, text=" [파일] ", bg="#f7f7f7", fg="#666").pack(side="left")
        tk.Button(toolbar_frame, text="💾 저장", command=self.controller.save_macro, relief=tk.FLAT, bg="#f7f7f7", cursor="hand2").pack(side="left", padx=2, pady=2)
        tk.Button(toolbar_frame, text="📂 DB가져오기", command=self.controller.load_macro_dialog, relief=tk.FLAT, bg="#f7f7f7", cursor="hand2").pack(side="left", padx=2, pady=2)

        ttk.Separator(toolbar_frame, orient="vertical").pack(side="left", fill="y", padx=5, pady=2)

        tk.Label(toolbar_frame, text="[편집] ", bg="#f7f7f7", fg="#666").pack(side="left")
        tk.Button(toolbar_frame, text="🗑️ 선택삭제", command=self.controller.delete_selected_macro, relief=tk.FLAT, bg="#f7f7f7", cursor="hand2").pack(side="left", padx=2, pady=2)
        tk.Button(toolbar_frame, text="🔄 초기화", command=self.controller.clear_macro, relief=tk.FLAT, bg="#f7f7f7", cursor="hand2").pack(side="left", padx=2, pady=2)

        ttk.Separator(toolbar_frame, orient="vertical").pack(side="left", fill="y", padx=5, pady=2)

        tk.Label(toolbar_frame, text="[보기] ", bg="#f7f7f7", fg="#666").pack(side="left")
        tk.Button(toolbar_frame, text="🔍+ 확대", command=lambda: self.controller.change_font_size(2), relief=tk.FLAT, bg="#f7f7f7", cursor="hand2").pack(side="left", padx=2, pady=2)
        tk.Button(toolbar_frame, text="🔍- 축소", command=lambda: self.controller.change_font_size(-2), relief=tk.FLAT, bg="#f7f7f7", cursor="hand2").pack(side="left", padx=2, pady=2)

        main_container = tk.Frame(self.root)
        main_container.pack(fill="both", expand=True, padx=10, pady=5)

        left_panel = tk.LabelFrame(main_container, text="저장된 매크로 리스트", padx=8, pady=8)
        left_panel.pack(side="left", fill="y", padx=(0, 5), pady=0)

        list_frame_left = tk.Frame(left_panel)
        list_frame_left.pack(fill="both", expand=True, pady=(0, 5))

        self.saved_listbox = tk.Listbox(list_frame_left, width=22, font=self.app_font)
        self.saved_listbox.pack(side="left", fill="both", expand=True)
        self.saved_listbox.bind("<Double-1>", lambda e: self.controller.load_selected_saved_macro())
        self.saved_listbox.bind("<MouseWheel>", lambda e: self.saved_listbox.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        sb_left = ttk.Scrollbar(list_frame_left, orient="vertical", command=self.saved_listbox.yview)
        sb_left.pack(side="right", fill="y")
        self.saved_listbox.configure(yscrollcommand=sb_left.set)

        tk.Button(left_panel, text="📂 불러오기", bg="#e6f2ff", command=self.controller.load_selected_saved_macro).pack(fill="x", pady=2)
        tk.Button(left_panel, text="✏️ 이름 수정", bg="#fff0b3", command=self.controller.rename_saved_macro).pack(fill="x", pady=2)
        tk.Button(left_panel, text="🗑️ 항목 삭제", bg="#ffcccc", command=self.controller.delete_saved_file).pack(fill="x", pady=2)
        tk.Button(left_panel, text="🔄 목록 새로고침", command=self.controller.refresh_saved_list).pack(fill="x", pady=2)

        right_panel = tk.Frame(main_container)
        right_panel.pack(side="right", fill="both", expand=True, padx=(5, 0), pady=0)

        top_frame = tk.LabelFrame(right_panel, text="기본 타겟 창 설정 및 고급 최적화 옵션", padx=10, pady=6)
        top_frame.pack(fill="x", pady=(0, 5))

        # 기본 창 제목 (상단)
        title_frame = tk.Frame(top_frame)
        title_frame.pack(fill="x", pady=(0, 8))
        tk.Label(title_frame, text="기본 창 제목:").pack(side="left", padx=2)
        self.target_window_entry = tk.Entry(title_frame, width=50)
        self.target_window_entry.pack(side="left", padx=2, fill="x", expand=True)
        
        # 옵션들 (하단)
        options_frame = tk.Frame(top_frame)
        options_frame.pack(fill="x", pady=(8, 0))
        tk.Checkbutton(options_frame, text="위치 자동 보정", variable=self.auto_correct_var).pack(side="left", padx=3)
        tk.Checkbutton(options_frame, text="크기 안정화", variable=self.dynamic_size_wait_var).pack(side="left", padx=3)
        tk.Checkbutton(options_frame, text="UIA 탐색", variable=self.uia_search_var).pack(side="left", padx=3)
        self.popup_checkbutton = tk.Checkbutton(
            options_frame,
            text="연동된 창(팝업) 적용",
            variable=self.popup_link_var,
            command=lambda: self.controller.ui.update_treeview(self.controller.engine.macro_list)
        )
        self.popup_checkbutton.pack(side="left", padx=3)

        loop_frame = tk.LabelFrame(right_panel, text="실행 반복 및 지연 설정", padx=10, pady=6)
        loop_frame.pack(fill="x", pady=(0, 5))

        tk.Label(loop_frame, text="반복 횟수:").pack(side="left", padx=2)
        self.loop_count_entry = tk.Entry(loop_frame, width=6)
        self.loop_count_entry.pack(side="left", padx=4)
        tk.Label(loop_frame, text="(0:무한)", fg="gray", font=("Malgun Gothic", 9)).pack(side="left", padx=2)

        tk.Label(loop_frame, text="회차간 대기(초):").pack(side="left", padx=8)
        self.loop_delay_entry = tk.Entry(loop_frame, width=6)
        self.loop_delay_entry.pack(side="left", padx=4)

        ctrl_frame = tk.Frame(right_panel, pady=2)
        ctrl_frame.pack(fill="x", pady=(0, 5))

        self.btn_record = tk.Button(ctrl_frame, text="🔴 기록 [F2]", bg="#ffcccb", command=self.controller.toggle_recording, font=("Malgun Gothic", self.current_font_size, "bold"), width=12, justify="center", anchor="center")
        self.btn_record.pack(side="left", padx=2)

        self.btn_run = tk.Button(ctrl_frame, text="▶️ 실행 [F3]", bg="#d0f0c0", command=self.controller.toggle_running, font=("Malgun Gothic", self.current_font_size, "bold"), width=12, justify="center", anchor="center")
        self.btn_run.pack(side="left", padx=2)

        self.btn_delete = tk.Button(ctrl_frame, text="🗑️ 선택삭제", bg="#ffebcc", command=self.controller.delete_selected_macro, font=("Malgun Gothic", self.current_font_size, "bold"), width=12, justify="center", anchor="center")
        self.btn_delete.pack(side="left", padx=2)

        self.btn_edit = tk.Button(ctrl_frame, text="✏️ 선택수정", bg="#e6f2ff", command=self.controller.edit_selected_macro_btn, font=("Malgun Gothic", self.current_font_size, "bold"), width=12, justify="center", anchor="center")
        self.btn_edit.pack(side="left", padx=2)

        self.btn_clear_all = tk.Button(ctrl_frame, text="🔄 초기화", bg="#ff9999", command=self.controller.clear_macro, font=("Malgun Gothic", self.current_font_size, "bold"), width=12, justify="center", anchor="center")
        self.btn_clear_all.pack(side="left", padx=2)

        list_frame_right = tk.LabelFrame(right_panel, text="매크로 동작 편집 (각 동작별 활성화 창/팝업창 자동 연동 표시)", padx=5, pady=5)
        list_frame_right.pack(fill="both", expand=True)

        columns = ("No", "동작", "X 좌표", "Y 좌표", "대기시간(초)", "연동된 창(팝업) 제목")
        self.tree = ttk.Treeview(list_frame_right, columns=columns, show="headings", height=8, selectmode="extended")
        
        saved_widths = self.controller.config.get("tree_column_widths", [45, 60, 75, 75, 95, 230])
        for idx, col in enumerate(columns):
            self.tree.heading(col, text=col)
            anchor_val = "w" if col == "연동된 창(팝업) 제목" else "center"
            width = saved_widths[idx] if idx < len(saved_widths) else 100
            self.tree.column(col, width=width, anchor=anchor_val)
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame_right, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.bind("<Double-1>", self.controller.edit_macro_item)

        self.status_var = tk.StringVar()
        self.status_var.set("준비 완료 (F2: 기록, F3: 실행/중단)")
        
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor="w", padx=5, pady=3, bg="#f0f0f0")
        status_bar.pack(side="bottom", fill="x", padx=10, pady=5)

    def apply_config(self, config):
        self.target_window_entry.delete(0, tk.END)
        self.target_window_entry.insert(0, config.get("target_window", ""))

        self.loop_count_entry.delete(0, tk.END)
        self.loop_count_entry.insert(0, config.get("loop_count", "1"))

        self.loop_delay_entry.delete(0, tk.END)
        self.loop_delay_entry.insert(0, config.get("loop_delay", "1.0"))

    def get_target_window(self):
        raw_title = self.target_window_entry.get().strip()
        if not hasattr(self.controller, "normalize_target_window_title"):
            return raw_title
        return self.controller.normalize_target_window_title(raw_title)

    def set_target_window(self, title):
        safe_title = title
        if hasattr(self.controller, "normalize_target_window_title"):
            safe_title = self.controller.normalize_target_window_title(title)
        self.target_window_entry.delete(0, tk.END)
        self.target_window_entry.insert(0, safe_title)
        
    def get_auto_correct(self):
        return self.auto_correct_var.get()

    def get_dynamic_size_wait(self):
        return self.dynamic_size_wait_var.get()

    def get_uia_search(self):
        return self.uia_search_var.get()

    def get_popup_link(self):
        return self.popup_link_var.get()

    def get_loop_count(self):
        return self.loop_count_entry.get().strip()
        
    def get_loop_delay(self):
        return self.loop_delay_entry.get().strip()

    def set_record_button_recording(self):
        self.btn_record.config(text="⏹️ 기록중지 [F2]", bg="#ff9999")

    def set_record_button_stopped(self):
        self.btn_record.config(text="🔴 기록 [F2]", bg="#ffcccb")
        
    def set_run_button_running(self):
        self.btn_run.config(text="⏹️ 중단 [F3]", bg="#ffb347")
        
    def set_run_button_stopped(self):
        self.btn_run.config(text="▶️ 실행 [F3]", bg="#d0f0c0")

    def update_treeview(self, macro_list):
        for row in self.tree.get_children():
            self.tree.delete(row)
        show_popup_title = self.get_popup_link()
        for idx, item in enumerate(macro_list):
            is_key_press = item[0] == "key_press"
            win_title = item[4] if len(item) >= 5 and not is_key_press else ""
            if not show_popup_title:
                win_title = ""
            try:
                formatted_delay = f"{float(item[3]):.2f}"
            except (ValueError, TypeError):
                formatted_delay = str(item[3])
            action_name = "엔터키" if is_key_press else item[0]
            x_value = "" if is_key_press else item[1]
            y_value = "" if is_key_press else item[2]
            self.tree.insert("", "end", values=(idx+1, action_name, x_value, y_value, formatted_delay, win_title))

        # 팝업 제목 표시 옵션에 따라 컬럼 제목을 현실적으로 반영
        if show_popup_title:
            self.tree.heading("연동된 창(팝업) 제목", text="연동된 창(팝업) 제목")
        else:
            self.tree.heading("연동된 창(팝업) 제목", text="연동된 창(팝업) 제목 (OFF)")

    def update_saved_listbox(self, db_keys):
        self.saved_listbox.delete(0, tk.END)
        for name in sorted(db_keys):
            self.saved_listbox.insert(tk.END, name)
            
    def get_selected_saved_macro(self):
        selected = self.saved_listbox.curselection()
        if not selected:
            return None
        return self.saved_listbox.get(selected[0])

    def get_selected_tree_indices(self):
        selected = self.tree.selection()
        return sorted([self.tree.index(item) for item in selected], reverse=True)
        
    def update_font_size(self, new_size):
        self.current_font_size = new_size
        self.app_font.configure(size=self.current_font_size)
        self.update_styles()
        
        ctrl_buttons = [self.btn_record, self.btn_run, self.btn_delete, self.btn_clear_all, self.btn_edit]
        for btn in ctrl_buttons:
            try:
                btn.config(font=("Malgun Gothic", self.current_font_size, "bold"))
            except Exception:
                pass
        self.create_menu()


# =========================================================================
# 4. 메인 컨트롤러 (MainController)
# =========================================================================
class MainController:
    def __init__(self, root):
        self.root = root
        self.root.title("스마트 마우스 매크로 프로그램 (Version 13f, 2026.10.19)")
        
        self.current_loaded_macro_name = None
        
        self.data_manager = DataManager()
        self.config = self.data_manager.load_config()
        self.root.geometry(self.config.get("geometry", "1050x650"))
        
        self.ui = MacroAppUI(self.root, self)
        self.engine = MacroEngine(self)

        self.config["target_window"] = ""
        self.ui.apply_config(self.config)
        self.refresh_saved_list()
        
        self.patch_messagebox()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def normalize_target_window_title(self, title):
        if title is None:
            return ""
        cleaned = str(title).strip()
        if not cleaned:
            return ""
        app_title = self.root.title().strip()
        if cleaned == app_title:
            return ""
        return cleaned

    def log_message(self, message):
        timestamp = time.strftime('%H:%M:%S')
        self.ui.status_var.set(f"[{timestamp}] {message}")

    def safe_log(self, message):
        self.root.after(0, lambda: self.log_message(message))

    def on_close(self):
        columns = ("No", "동작", "X 좌표", "Y 좌표", "대기시간(초)", "연동된 창(팝업) 제목")
        widths = [self.ui.tree.column(col, "width") for col in columns]
        self.config["tree_column_widths"] = widths

        self.config["target_window"] = self.normalize_target_window_title(self.ui.get_target_window())
        self.config["loop_count"] = self.ui.get_loop_count()
        self.config["loop_delay"] = self.ui.get_loop_delay()
        self.config["font_size"] = self.ui.current_font_size
        self.config["geometry"] = self.root.geometry()
        self.config["auto_correct"] = self.ui.get_auto_correct()
        self.config["dynamic_size_wait"] = self.ui.get_dynamic_size_wait()
        self.config["uia_search"] = self.ui.get_uia_search()
        self.config["popup_link"] = self.ui.get_popup_link()
        self.data_manager.save_config(self.config)
        self.root.destroy()

    def change_font_size(self, delta):
        new_size = max(8, min(24, self.ui.current_font_size + delta))
        self.ui.update_font_size(new_size)
        
        current_geom = self.root.geometry().split("+")[0]
        try:
            w, h = map(int, current_geom.split("x"))
            new_w = max(1050, w + (delta * 40))
            new_h = max(650, h + (delta * 25))
            self.root.geometry(f"{new_w}x{new_h}")
        except Exception:
            pass
        
        self.safe_log(f"글자 크기가 {new_size}(으)로 변경되었습니다.")

    def show_help_dialog(self):
        help_win = tk.Toplevel(self.root)
        help_win.title("프로그램 도움말")
        
        # 저장된 도움말 창 크기/위치 불러오기 (기본값: 600x550)
        saved_geom = self.config.get("help_win_geom", "600x550")
        help_win.geometry(saved_geom)
        help_win.transient(self.root)
        help_win.grab_set()

        text_area = tk.Text(help_win, wrap="word", font=("Malgun Gothic", self.ui.current_font_size), padx=10, pady=10)
        text_area.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(text_area, command=text_area.yview)
        scrollbar.pack(side="right", fill="y")
        text_area.configure(yscrollcommand=scrollbar.set)

        help_text = """[ 스마트 마우스 매크로 도움말 ]

1. 매크로 기록 (F2)
- F2 키를 누르거나 [🔴 기록] 버튼을 클릭하면 기록이 시작됩니다.
- 대상 창이나 화면의 특정 위치를 클릭하면 해당 동작이 기록됩니다.
- 기록 시 대상 창의 제목을 자동으로 감지하여 저장합니다.
- 다시 F2를 누르면 기록이 중지됩니다.

2. 매크로 실행 (F3)
- F3 키를 누르거나 [▶️ 실행] 버튼을 클릭하면 기록된 동작이 실행됩니다.
- 지정된 반복 횟수와 대기 시간에 따라 매크로가 작동합니다.
- 실행 중 다시 F3을 누르면 즉시 중단됩니다.

3. 편집 및 초기화
- 선택 삭제: 목록에서 선택한 동작만 삭제합니다.
- 초기화: 기록된 전체 매크로 동작과 기본 대상 창 제목을 지웁니다.
- 항목을 더블클릭하면 해당 동작의 세부 설정(좌표, 딜레이, 대상 창)을 수정할 수 있습니다.

4. 파일 저장 및 불러오기
- 작성한 매크로를 데이터베이스에 새 이름으로 저장할 수 있습니다.
- 좌측 목록에서 저장된 매크로를 더블클릭하거나 [📂 불러오기]를 눌러 사용할 수 있습니다.

5. 고급 최적화 옵션
- 위치 자동 보정: 창이 이동했을 때 현재 창 위치를 기준으로 클릭 좌표를 자동 계산합니다.
- 크기 안정화: 대상 창의 크기가 변하고 있을 때 멈출 때까지 잠시 대기합니다.
- UIA 탐색: 윈도우 UI 자동화(UIA)를 사용하여 대상 창을 정밀하게 탐색합니다.
- 연동된 창(팝업) 적용: 동작별로 감지된 개별 팝업창을 기준으로 동작을 실행할지 선택합니다. (해제 시 기본 창 기준으로만 동작)
"""
        text_area.insert("1.0", help_text)
        text_area.configure(state="disabled")

        # 닫기 시 현재 창 크기/위치 저장
        def on_help_close():
            self.config["help_win_geom"] = help_win.geometry()
            self.data_manager.save_config(self.config)
            help_win.destroy()

        help_win.protocol("WM_DELETE_WINDOW", on_help_close)
        tk.Button(help_win, text="닫기", command=on_help_close, font=("Malgun Gothic", self.ui.current_font_size), bg="#e6f2ff", width=10).pack(pady=(0, 10))

    def patch_messagebox(self):
        def custom_askyesno(title, message, **kwargs):
            return self.custom_messagebox(title, message, m_type="question")
        
        def custom_showwarning(title, message, **kwargs):
            return self.custom_messagebox(title, message, m_type="warning")
            
        def custom_showerror(title, message, **kwargs):
            return self.custom_messagebox(title, message, m_type="error")
            
        def custom_showinfo(title, message, **kwargs):
            return self.custom_messagebox(title, message, m_type="info")

        messagebox.askyesno = custom_askyesno
        messagebox.showwarning = custom_showwarning
        messagebox.showerror = custom_showerror
        messagebox.showinfo = custom_showinfo

    def custom_messagebox(self, title, message, m_type="info", config_key="msgbox_geom", default_geom="450x250"):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        
        saved_geom = self.config.get(config_key, default_geom)
        dialog.geometry(saved_geom)
        dialog.minsize(450, 250)
        dialog.transient(self.root)
        dialog.grab_set()

        result = [None]
        
        frame = tk.Frame(dialog, padx=15, pady=15)
        frame.pack(fill="both", expand=True)

        lbl = tk.Label(frame, text=message, font=("Malgun Gothic", self.ui.current_font_size), justify="left", anchor="nw")
        lbl.pack(fill="both", expand=True, pady=(0, 15))

        btn_frame = tk.Frame(frame)
        btn_frame.pack(side="bottom", pady=(0, 10))

        def set_result(val):
            result[0] = val
            self.config[config_key] = dialog.geometry()
            self.data_manager.save_config(self.config)
            dialog.destroy()

        if m_type == "question":
            tk.Button(btn_frame, text="예", command=lambda: set_result(True), width=8, bg="#d0f0c0", font=("Malgun Gothic", self.ui.current_font_size)).pack(side="left", padx=5)
            tk.Button(btn_frame, text="아니요", command=lambda: set_result(False), width=8, font=("Malgun Gothic", self.ui.current_font_size)).pack(side="left", padx=5)
            dialog.bind("<Return>", lambda e: set_result(True))
            dialog.bind("<Escape>", lambda e: set_result(False))
            dialog.protocol("WM_DELETE_WINDOW", lambda: set_result(False))
        else:
            tk.Button(btn_frame, text="확인", command=lambda: set_result(True), width=8, bg="#d0f0c0", font=("Malgun Gothic", self.ui.current_font_size)).pack(side="left", padx=5)
            dialog.bind("<Return>", lambda e: set_result(True))
            dialog.bind("<Escape>", lambda e: set_result(True))
            dialog.protocol("WM_DELETE_WINDOW", lambda: set_result(True))

        self.root.wait_window(dialog)
        return result[0]

    def custom_askstring(self, title, prompt, initialvalue="", config_key="askstring_geom", default_geom="400x200"):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        
        saved_geom = self.config.get(config_key, default_geom)
        dialog.geometry(saved_geom)
        dialog.minsize(400, 200)
        dialog.transient(self.root)
        dialog.grab_set()

        result: list[str | None] = [None]
        
        frame = tk.Frame(dialog, padx=15, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=prompt, font=("Malgun Gothic", self.ui.current_font_size)).pack(anchor="w", pady=(0, 10))
        entry = tk.Entry(frame, width=30, font=("Malgun Gothic", self.ui.current_font_size))
        if initialvalue:
            entry.insert(0, initialvalue)
        entry.pack(fill="x", pady=(0, 15))
        entry.focus_set()

        btn_frame = tk.Frame(frame)
        btn_frame.pack(side="bottom", pady=(0, 10))

        def on_ok(event=None):
            result[0] = entry.get()
            self.config[config_key] = dialog.geometry()
            self.data_manager.save_config(self.config)
            dialog.destroy()

        def on_cancel(event=None):
            self.config[config_key] = dialog.geometry()
            self.data_manager.save_config(self.config)
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        dialog.bind("<Return>", on_ok)
        dialog.bind("<Escape>", on_cancel)

        tk.Button(btn_frame, text="확인", command=on_ok, width=8, bg="#d0f0c0", font=("Malgun Gothic", self.ui.current_font_size)).pack(side="left", padx=5)
        tk.Button(btn_frame, text="취소", command=on_cancel, width=8, font=("Malgun Gothic", self.ui.current_font_size)).pack(side="left", padx=5)

        self.root.wait_window(dialog)
        return result[0]

    def toggle_recording(self):
        is_recording = self.engine.toggle_recording()
        if is_recording:
            self.current_loaded_macro_name = None
            self.ui.set_record_button_recording()
        else:
            self.ui.set_record_button_stopped()

    def toggle_running(self):
        is_running = self.engine.toggle_running()
        if is_running:
            self.ui.set_run_button_running()
        else:
            self.ui.set_run_button_stopped()

    def delete_selected_macro(self):
        indices = self.ui.get_selected_tree_indices()
        if not indices:
            messagebox.showwarning("경고", "삭제할 동작 항목을 선택해주세요.")
            return
        
        for idx in indices:
            if 0 <= idx < len(self.engine.macro_list):
                del self.engine.macro_list[idx]
                
        self.ui.update_treeview(self.engine.macro_list)
        self.safe_log(f"선택한 동작 항목 {len(indices)}개가 삭제되었습니다.")

    def clear_macro(self):
        if not self.engine.macro_list and not self.config.get("target_window"):
            # 데이터가 없으면 확인 없이 진행 (또는 이미 초기화된 상태)
            pass
        else:
            result = messagebox.askyesno(
                "초기화 확인",
                "모든 매크로 동작 기록과 기본 창 제목을 초기화하시겠습니까?",
                icon='warning'
            )
            if not result:
                return

        self.engine.macro_list.clear()
        self.current_loaded_macro_name = None
        self.engine.first_click_window_title = None
        self.ui.set_target_window("")
        self.config["target_window"] = ""
        self.ui.update_treeview(self.engine.macro_list)
        self.safe_log("모든 매크로 기록과 기본 창 제목이 초기화되었습니다.")

    def refresh_saved_list(self):
        db = self.data_manager.load_macro_db()
        self.ui.update_saved_listbox(db.keys())

    def load_selected_saved_macro(self):
        name = self.ui.get_selected_saved_macro()
        if not name:
            messagebox.showwarning("경고", "불러올 매크로 항목을 선택해주세요.")
            return
        
        db = self.data_manager.load_macro_db()
        if name in db:
            saved_data = db[name]
            
            # 새 형식 (딕셔너리): {"target_window": "...", "macro_list": [...]}
            if isinstance(saved_data, dict):
                target_window = saved_data.get("target_window", "")
                self.engine.macro_list = saved_data.get("macro_list", [])
            # 기존 형식 (리스트): [...] - 하위 호환성
            else:
                target_window = ""
                self.engine.macro_list = saved_data
            
            # 매크로 리스트 정규화
            for i, item in enumerate(self.engine.macro_list):
                if len(item) < 5:
                    self.engine.macro_list[i] = (item[0], item[1], item[2], item[3], "")
                elif len(item) > 5:
                    self.engine.macro_list[i] = item[:5]
            
            # 기본 창 제목 복원
            if target_window:
                self.ui.set_target_window(target_window)
                self.config["target_window"] = target_window
            
            self.ui.update_treeview(self.engine.macro_list)
            self.current_loaded_macro_name = name
            self.safe_log(f"매크로 불러오기 성공: {name}")
        else:
            messagebox.showerror("에러", "해당 매크로 정보를 찾을 수 없습니다.")

    def rename_saved_macro(self):
        old_name = self.ui.get_selected_saved_macro()
        if not old_name:
            messagebox.showwarning("경고", "이름을 수정할 매크로 항목을 선택해주세요.")
            return
        
        new_name = self.custom_askstring("매크로 이름 수정", "새로운 이름을 입력하세요:", initialvalue=old_name, config_key="askstring_rename_geom")
        if new_name:
            new_name = new_name.strip()
            if not new_name or new_name == old_name:
                return
            
            db = self.data_manager.load_macro_db()
            if new_name in db:
                messagebox.showwarning("경고", "이미 존재하는 매크로 이름입니다.")
                return
            
            if old_name in db:
                db[new_name] = db.pop(old_name)
                self.data_manager.save_macro_db(db)
                if self.current_loaded_macro_name == old_name:
                    self.current_loaded_macro_name = new_name
                self.refresh_saved_list()
                self.safe_log(f"매크로 이름이 '{new_name}'(으)로 변경되었습니다.")

    def delete_saved_file(self):
        name = self.ui.get_selected_saved_macro()
        if not name:
            messagebox.showwarning("경고", "삭제할 매크로 항목을 선택해주세요.")
            return
        
        if messagebox.askyesno("확인", f"'{name}' 매크로 항목을 영구 삭제하시겠습니까?"):
            db = self.data_manager.load_macro_db()
            if name in db:
                del db[name]
                if self.current_loaded_macro_name == name:
                    self.current_loaded_macro_name = None
                    self.engine.macro_list.clear()
                    self.ui.update_treeview(self.engine.macro_list)
                self.data_manager.save_macro_db(db)
                self.refresh_saved_list()
                self.safe_log(f"매크로 '{name}'이 삭제되었습니다.")

    def save_macro(self):
        if not self.engine.macro_list:
            messagebox.showwarning("경고", "저장할 매크로 기록이 없습니다.")
            return
        
        default_name = self.current_loaded_macro_name if self.current_loaded_macro_name else ""
        
        macro_name = self.custom_askstring("매크로 저장", "저장할 매크로 이름을 입력하세요:", initialvalue=default_name, config_key="askstring_save_geom")
        if macro_name:
            macro_name = macro_name.strip()
            if not macro_name:
                return
            
            # 기본 창 제목과 매크로 리스트를 함께 저장
            target_window = self.ui.get_target_window()
            macro_data = {
                "target_window": target_window,
                "macro_list": self.engine.macro_list
            }
            
            db = self.data_manager.load_macro_db()
            db[macro_name] = macro_data
            self.data_manager.save_macro_db(db)
            self.current_loaded_macro_name = macro_name
            self.refresh_saved_list()
            self.safe_log(f"매크로 저장 완료: {macro_name} (기본 창: '{target_window}')")

    def load_macro_dialog(self):
        file_path = filedialog.askopenfilename(defaultextension=".json", filetypes=[("JSON 파일", "*.json")])
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        for m_name, m_data in data.items():
                            # 새 형식 (딕셔너리)인 경우
                            if isinstance(m_data, dict):
                                macro_list = m_data.get("macro_list", [])
                            # 기존 형식 (리스트)인 경우
                            else:
                                macro_list = m_data
                            
                            # 매크로 리스트 정규화
                            for i, item in enumerate(macro_list):
                                if len(item) < 5:
                                    macro_list[i] = (item[0], item[1], item[2], item[3], "")
                                elif len(item) > 5:
                                    macro_list[i] = item[:5]
                        
                        db = self.data_manager.load_macro_db()
                        db.update(data)
                        self.data_manager.save_macro_db(db)
                        self.refresh_saved_list()
                        self.safe_log(f"매크로 DB 파일 병합 완료: {os.path.basename(file_path)}")
                    else:
                        messagebox.showerror("에러", "올바른 매크로 DB 형식이 아닙니다.")
            except Exception as e:
                messagebox.showerror("에러", f"파일을 불러오지 못했습니다: {e}")

    def edit_selected_macro_btn(self):
        indices = self.ui.get_selected_tree_indices()
        if not indices:
            messagebox.showwarning("경고", "수정할 매크로 항목을 선택해주세요.")
            return
        self.open_edit_window(indices[0])

    def edit_macro_item(self, event):
        indices = self.ui.get_selected_tree_indices()
        if not indices:
            return
        self.open_edit_window(indices[0])

    def open_edit_window(self, idx):
        current_data = self.engine.macro_list[idx]

        edit_win = tk.Toplevel(self.root)
        edit_win.title(f"매크로 기록 [{idx+1}] 상세 편집")
        
        saved_edit_geom = self.config.get("edit_win_geom", "380x400")
        edit_win.geometry(saved_edit_geom)
        edit_win.resizable(True, True)

        frame = tk.Frame(edit_win, padx=15, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="X 상대좌표:", font=self.ui.app_font).pack(anchor="w", pady=(5, 2))
        ent_x = tk.Entry(frame, width=30, font=self.ui.app_font)
        ent_x.insert(0, str(current_data[1]))
        ent_x.pack(fill="x", pady=(0, 10))

        tk.Label(frame, text="Y 상대좌표:", font=self.ui.app_font).pack(anchor="w", pady=(2, 2))
        ent_y = tk.Entry(frame, width=30, font=self.ui.app_font)
        ent_y.insert(0, str(current_data[2]))
        ent_y.pack(fill="x", pady=(0, 10))

        tk.Label(frame, text="대기시간(초):", font=self.ui.app_font).pack(anchor="w", pady=(2, 2))
        ent_delay = tk.Entry(frame, width=30, font=self.ui.app_font)
        ent_delay.insert(0, str(current_data[3]))
        ent_delay.pack(fill="x", pady=(0, 10))

        tk.Label(frame, text="연동된 창(팝업) 제목:", font=self.ui.app_font).pack(anchor="w", pady=(2, 2))
        ent_wintitle = tk.Entry(frame, width=30, font=self.ui.app_font)
        if self.ui.get_popup_link():
            ent_wintitle.insert(0, str(current_data[4]) if len(current_data) >= 5 else "")
        else:
            ent_wintitle.insert(0, "")
        ent_wintitle.pack(fill="x", pady=(0, 15))

        def save_edit():
            try:
                nd = float(ent_delay.get())
                action_type = current_data[0] if current_data else "click"

                if action_type == "key_press":
                    nx, ny = current_data[1], current_data[2]
                    nw = current_data[4] if len(current_data) >= 5 else ""
                else:
                    nx = int(ent_x.get())
                    ny = int(ent_y.get())
                    nw = ent_wintitle.get().strip() if self.ui.get_popup_link() else ""
                
                self.config["edit_win_geom"] = edit_win.geometry()
                self.data_manager.save_config(self.config)

                self.engine.macro_list[idx] = (action_type, nx, ny, nd, nw)
                self.ui.update_treeview(self.engine.macro_list)
                self.safe_log(f"항목 {idx+1} 수정 완료")
                edit_win.destroy()
            except ValueError:
                messagebox.showerror("에러", "올바른 숫자 형식을 입력해주세요.")

        def on_edit_close():
            try:
                self.config["edit_win_geom"] = edit_win.geometry()
                self.data_manager.save_config(self.config)
            except Exception:
                pass
            edit_win.destroy()

        edit_win.protocol("WM_DELETE_WINDOW", on_edit_close)
        tk.Button(frame, text="수정 내용 저장", command=save_edit, bg="#d0f0c0", font=self.ui.app_font, height=1).pack(fill="x", pady=(5, 0))

if __name__ == "__main__":
    root = tk.Tk()
    app = MainController(root)
    root.mainloop()