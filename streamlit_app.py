#Cursor가 다듬어줌. 2025.08.13

import time
from datetime import datetime, timedelta
import json
import os
import threading
import logging
import sys

# GUI/automation libraries that require a display will be imported lazily
# to avoid import-time failures on headless hosts (e.g., Streamlit Cloud).
tk = None
ttk = None
simpledialog = None
messagebox = None
filedialog = None
Calendar = None
Image = None
ImageTk = None
pyautogui = None
keyboard = None

# Helper flag: True when GUI automation is available
AUTOMATION_AVAILABLE = False

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tnlcopy.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class TNLMacro:
    def __init__(self):
        self.config_path = os.path.join(os.path.expanduser("~"), "Desktop", "tnlcopy_config.json")
        self.coords = {}
        self.is_running = False
        self.is_paused = False
        self.current_date_index = 0
        self.total_dates = 0
        self.dates = []
        self.ref_date = ""
        # Try to initialize optional libraries (lazy, safe)
        self._init_optional_libs()

        # If Tkinter wasn't loaded (e.g., running on Streamlit/cloud), skip GUI setup
        if tk is not None:
            # PyAutoGUI 설정 if available
            if pyautogui is not None:
                try:
                    pyautogui.FAILSAFE = True
                    pyautogui.PAUSE = 0.5
                except Exception:
                    # ignore pyautogui configuration errors at init
                    pass

            self.setup_gui()
        else:
            # Running in headless mode (Streamlit). Create minimal placeholders
            logging.info("Running in headless mode: GUI setup skipped")

    def _init_optional_libs(self):
        """Try to import optional GUI and automation libraries safely.
        Sets module-level globals when imports succeed. This avoids
        import-time failures on headless hosts (like Streamlit Cloud).
        """
        global tk, ttk, simpledialog, messagebox, filedialog
        global Calendar, Image, ImageTk, pyautogui, keyboard, AUTOMATION_AVAILABLE

        # Try tkinter (may fail on headless linux)
        try:
            import tkinter as _tk
            from tkinter import ttk as _ttk, simpledialog as _sd, messagebox as _mb, filedialog as _fd
            tk = _tk
            ttk = _ttk
            simpledialog = _sd
            messagebox = _mb
            filedialog = _fd
        except Exception:
            tk = None
            ttk = None
            simpledialog = None
            messagebox = None
            filedialog = None

        # tkcalendar (optional)
        try:
            from tkcalendar import Calendar as _Calendar
            Calendar = _Calendar
        except Exception:
            Calendar = None

        # PIL (optional)
        try:
            from PIL import Image as _Image, ImageTk as _ImageTk
            Image = _Image
            ImageTk = _ImageTk
        except Exception:
            Image = None
            ImageTk = None

        # GUI automation libs: only attempt when display is likely available
        AUTOMATION_AVAILABLE = False
        try:
            should_try = (os.name == 'nt') or ('DISPLAY' in os.environ)
            if should_try:
                import pyautogui as _pyautogui
                import keyboard as _keyboard
                pyautogui = _pyautogui
                keyboard = _keyboard
                AUTOMATION_AVAILABLE = True
        except Exception:
            pyautogui = None
            keyboard = None
            AUTOMATION_AVAILABLE = False
    
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("T&L 자동 복사 매크로 v2.0")
        self.root.geometry("600x700")
        self.root.resizable(True, True)
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="T&L 자동 복사 매크로", 
                               font=("맑은 고딕", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 설정 섹션
        self.create_settings_section(main_frame)
        
        # 진행률 섹션
        self.create_progress_section(main_frame)
        
        # 로그 섹션
        self.create_log_section(main_frame)
        
        # 버튼 섹션
        self.create_button_section(main_frame)
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 초기 설정 로드
        self.load_config()
        
        # ESC 키 감지 스레드
        self.esc_thread = threading.Thread(target=self.monitor_esc_key, daemon=True)
        self.esc_thread.start()
    
    def create_settings_section(self, parent):
        # 설정 프레임
        settings_frame = ttk.LabelFrame(parent, text="설정", padding="10")
        settings_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 좌표 설정 버튼
        ttk.Button(settings_frame, text="좌표 설정", 
                  command=self.set_coords_gui).grid(row=0, column=0, padx=(0, 10))
        
        # 설정 저장/불러오기
        ttk.Button(settings_frame, text="설정 저장", 
                  command=self.save_config).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(settings_frame, text="설정 불러오기", 
                  command=self.load_config).grid(row=0, column=2)
        
        # 날짜 설정
        date_frame = ttk.Frame(settings_frame)
        date_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(date_frame, text="작업 날짜:").grid(row=0, column=0, sticky=tk.W)
        self.date_label = ttk.Label(date_frame, text="선택되지 않음", foreground="red")
        self.date_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        ttk.Button(date_frame, text="날짜 선택", 
                  command=self.select_dates).grid(row=0, column=2, padx=(10, 0))
        
        ttk.Label(date_frame, text="복사 기준일자:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.ref_date_label = ttk.Label(date_frame, text="선택되지 않음", foreground="red")
        self.ref_date_label.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(5, 0))
        ttk.Button(date_frame, text="기준일자 선택", 
                  command=self.select_ref_date).grid(row=1, column=2, padx=(10, 0), pady=(5, 0))
    
    def create_progress_section(self, parent):
        # 진행률 프레임
        progress_frame = ttk.LabelFrame(parent, text="진행 상황", padding="10")
        progress_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 진행률 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                          maximum=100, length=400)
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 진행률 텍스트
        self.progress_text = ttk.Label(progress_frame, text="대기 중...")
        self.progress_text.grid(row=1, column=0, columnspan=2, sticky=tk.W)
        
        # 현재 작업 표시
        self.current_work_label = ttk.Label(progress_frame, text="")
        self.current_work_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
    
    def create_log_section(self, parent):
        # 로그 프레임
        log_frame = ttk.LabelFrame(parent, text="작업 로그", padding="10")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 로그 텍스트 영역
        log_container = ttk.Frame(log_frame)
        log_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.log_text = tk.Text(log_container, height=8, width=70, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_container, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
    
    def create_button_section(self, parent):
        # 버튼 프레임
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(0, 10))
        
        # 시작/일시정지/중지 버튼
        self.start_button = ttk.Button(button_frame, text="작업 시작", 
                                      command=self.start_work, style="Accent.TButton")
        self.start_button.grid(row=0, column=0, padx=(0, 10))
        
        self.pause_button = ttk.Button(button_frame, text="일시정지", 
                                      command=self.pause_work, state="disabled")
        self.pause_button.grid(row=0, column=1, padx=(0, 10))
        
        self.stop_button = ttk.Button(button_frame, text="작업 중지", 
                                     command=self.stop_work, state="disabled")
        self.stop_button.grid(row=0, column=2)
        
        # 상태 표시
        self.status_label = ttk.Label(button_frame, text="대기 중", 
                                     font=("맑은 고딕", 10, "bold"))
        self.status_label.grid(row=1, column=0, columnspan=3, pady=(10, 0))
    
    def log_message(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        
        if level == "ERROR":
            logging.error(message)
        elif level == "WARNING":
            logging.warning(message)
        else:
            logging.info(message)
    
    def set_coords_gui(self):
        if tk is None or pyautogui is None:
            # Running headless or automation libs not available
            logging.warning("set_coords_gui called but GUI/automation is not available")
            if messagebox is not None:
                try:
                    messagebox.showwarning("사용 불가", "좌표 설정은 GUI 환경(로컬 Windows 등)에서만 동작합니다.")
                except Exception:
                    pass
            return

        if not self.coords:
            self.coords = {}
        
        coord_names = [
            ("date_xy", "일자 입력란"),
            ("serch_xy", "조회 버튼"),
            ("ref_date_xy", "복사 기준일자 입력란"),
            ("select_xy", "사원 선택란"),
            ("ref_copy_xy", "이전실적복사 버튼"),
            ("copy_xy", "복사 버튼")
        ]
        
        for key, label in coord_names:
            if not self.get_coordinate(label):
                break
        else:
            messagebox.showinfo("완료", "모든 좌표가 설정되었습니다.")
            self.save_config()
            return
        
        self.log_message("좌표 설정이 완료되지 않았습니다. 모든 좌표를 설정해주세요.")
    
    def get_coordinate(self, label):
        popup = tk.Toplevel(self.root)
        popup.title(f"좌표 설정 - {label}")
        popup.geometry("500x400")
        popup.grab_set()
        
        # 화면 중앙 배치
        popup.update_idletasks()
        w = popup.winfo_width()
        h = popup.winfo_height()
        sw = popup.winfo_screenwidth()
        sh = popup.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        popup.geometry(f"+{x}+{y}")
        
        frame = ttk.Frame(popup, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"{label} 위치를 설정합니다.", 
                 font=("맑은 고딕", 12, "bold")).pack(pady=(0, 10))
        
        ttk.Label(frame, text="1. OK 버튼을 클릭하세요\n2. 3초 이내에 해당 위치에 마우스를 이동하세요\n3. 자동으로 좌표가 기록됩니다", 
                 font=("맑은 고딕", 10)).pack(pady=(0, 20))
        
        # 이미지 표시 (있는 경우)
        try:
            img_filename = f"{label}.png"
            if hasattr(sys, '_MEIPASS'):
                img_path = os.path.join(sys._MEIPASS, img_filename)
            else:
                img_path = os.path.join(os.path.dirname(__file__), img_filename)
            
            if os.path.exists(img_path):
                img = Image.open(img_path)
                img.thumbnail((300, 300))
                tk_img = ImageTk.PhotoImage(img)
                img_label = ttk.Label(frame, image=tk_img)
                img_label.image = tk_img
                img_label.pack(pady=(0, 20))
        except Exception as e:
            self.log_message(f"이미지 로드 실패: {e}", "WARNING")
        
        coord_result = [None]
        
        def capture_coord():
            popup.destroy()
            time.sleep(3)
            x, y = pyautogui.position()
            coord_result[0] = (x, y)
            self.coords[label] = (x, y)
            self.log_message(f"{label} 좌표 설정: ({x}, {y})")
        
        ttk.Button(frame, text="OK", command=capture_coord, 
                  style="Accent.TButton").pack()
        
        popup.wait_window()
        return coord_result[0] is not None
    
    def select_dates(self):
        win = tk.Toplevel(self.root)
        win.title("작업 날짜 선택")
        win.geometry("400x500")
        win.grab_set()
        
        frame = ttk.Frame(win, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="작업할 날짜를 선택하세요", 
                 font=("맑은 고딕", 12, "bold")).pack(pady=(0, 10))
        
        cal = Calendar(frame, selectmode='day', date_pattern='yyyy-mm-dd', 
                      font=("맑은 고딕", 12))
        cal.pack(pady=(0, 10))
        
        selected_dates = []
        listbox = tk.Listbox(frame, selectmode=tk.EXTENDED, width=20, height=8, 
                            font=("맑은 고딕", 10))
        listbox.pack(pady=(0, 10))
        
        def add_date():
            date = cal.get_date()
            if date not in selected_dates:
                selected_dates.append(date)
                listbox.insert(tk.END, date)
        
        def remove_date():
            selection = listbox.curselection()
            for idx in reversed(selection):
                selected_dates.pop(idx)
                listbox.delete(idx)
        
        def clear_all():
            selected_dates.clear()
            listbox.delete(0, tk.END)
        
        def done():
            self.dates = selected_dates.copy()
            self.total_dates = len(self.dates)
            if self.dates:
                self.date_label.config(text=f"{len(self.dates)}개 날짜 선택됨", foreground="green")
                self.log_message(f"작업 날짜 설정: {len(self.dates)}개")
            win.destroy()
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        
        ttk.Button(btn_frame, text="추가", command=add_date).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="삭제", command=remove_date).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="전체 삭제", command=clear_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="완료", command=done, style="Accent.TButton").pack(side=tk.LEFT, padx=2)
    
    def select_ref_date(self):
        win = tk.Toplevel(self.root)
        win.title("복사 기준일자 선택")
        win.geometry("300x400")
        win.grab_set()
        
        frame = ttk.Frame(win, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="복사 기준일자를 선택하세요", 
                 font=("맑은 고딕", 12, "bold")).pack(pady=(0, 10))
        
        cal = Calendar(frame, selectmode='day', date_pattern='yyyy-mm-dd', 
                      font=("맑은 고딕", 12))
        cal.pack(pady=(0, 20))
        
        def done():
            self.ref_date = cal.get_date()
            self.ref_date_label.config(text=self.ref_date, foreground="green")
            self.log_message(f"복사 기준일자 설정: {self.ref_date}")
            win.destroy()
        
        ttk.Button(frame, text="선택", command=done, style="Accent.TButton").pack()
    
    def save_config(self):
        config = {
            'coords': self.coords,
            'last_ref_date': self.ref_date,
            'last_dates': self.dates
        }
        
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.log_message("설정이 저장되었습니다.")
        except Exception as e:
            self.log_message(f"설정 저장 실패: {e}", "ERROR")
    
    def load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                
                self.coords = config.get('coords', {})
                self.ref_date = config.get('last_ref_date', '')
                self.dates = config.get('last_dates', [])
                
                if self.ref_date:
                    self.ref_date_label.config(text=self.ref_date, foreground="green")
                if self.dates:
                    self.date_label.config(text=f"{len(self.dates)}개 날짜 선택됨", foreground="green")
                    self.total_dates = len(self.dates)
                
                self.log_message("설정을 불러왔습니다.")
            else:
                self.log_message("저장된 설정이 없습니다.")
        except Exception as e:
            self.log_message(f"설정 불러오기 실패: {e}", "ERROR")
    
    def start_work(self):
        if not self.validate_settings():
            return
        
        self.is_running = True
        self.is_paused = False
        self.current_date_index = 0
        
        self.start_button.config(state="disabled")
        self.pause_button.config(state="normal")
        self.stop_button.config(state="normal")
        self.status_label.config(text="작업 중...", foreground="blue")
        
        self.log_message("작업을 시작합니다.")
        
        # 작업 스레드 시작
        work_thread = threading.Thread(target=self.work_process, daemon=True)
        work_thread.start()
    
    def pause_work(self):
        if self.is_paused:
            self.is_paused = False
            self.pause_button.config(text="일시정지")
            self.status_label.config(text="작업 중...", foreground="blue")
            self.log_message("작업을 재개합니다.")
        else:
            self.is_paused = True
            self.pause_button.config(text="재개")
            self.status_label.config(text="일시정지", foreground="orange")
            self.log_message("작업을 일시정지합니다.")
    
    def stop_work(self):
        self.is_running = False
        self.is_paused = False
        
        self.start_button.config(state="normal")
        self.pause_button.config(state="disabled")
        self.stop_button.config(state="disabled")
        self.status_label.config(text="중지됨", foreground="red")
        
        self.log_message("작업이 중지되었습니다.")
    
    def validate_settings(self):
        if not self.coords:
            messagebox.showerror("오류", "좌표가 설정되지 않았습니다. 좌표를 먼저 설정하세요.")
            return False
        
        if not self.dates:
            messagebox.showerror("오류", "작업 날짜가 선택되지 않았습니다.")
            return False
        
        if not self.ref_date:
            messagebox.showerror("오류", "복사 기준일자가 선택되지 않았습니다.")
            return False
        
        return True
    
    def work_process(self):
        try:
            # 복사 기준일자 입력 (한 번만)
            self.log_message("복사 기준일자를 입력합니다...")
            self.current_work_label.config(text="복사 기준일자 입력 중...")
            
            pyautogui.click(*self.coords["복사 기준일자 입력란"])
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.press('delete')
            time.sleep(0.5)
            for char in self.ref_date:
                pyautogui.press(char)
                time.sleep(0.1)
            time.sleep(1)
            
            # 날짜별 반복 작업
            for i, date in enumerate(self.dates):
                if not self.is_running:
                    break
                
                while self.is_paused:
                    time.sleep(0.1)
                    if not self.is_running:
                        break
                
                self.current_date_index = i
                progress = ((i + 1) / len(self.dates)) * 100
                self.progress_var.set(progress)
                self.progress_text.config(text=f"{i + 1} / {len(self.dates)} ({progress:.1f}%)")
                self.current_work_label.config(text=f"처리 중: {date}")
                
                self.log_message(f"날짜 {date} 처리 시작 ({i + 1}/{len(self.dates)})")
                
                # 날짜 입력
                pyautogui.click(*self.coords["일자 입력란"])
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.5)
                pyautogui.press('delete')
                time.sleep(0.5)
                for char in date:
                    pyautogui.press(char)
                    time.sleep(0.1)
                time.sleep(1)
                
                # 조회
                pyautogui.click(*self.coords["조회 버튼"])
                time.sleep(2)
                
                # 사원 선택
                pyautogui.click(*self.coords["사원 선택란"])
                time.sleep(1)
                
                # 이전실적복사
                pyautogui.click(*self.coords["이전실적복사 버튼"])
                time.sleep(1)
                
                # 팝업 처리
                pyautogui.press('enter')
                time.sleep(1)
                
                # 복사
                pyautogui.click(*self.coords["복사 버튼"])
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(1)
                
                self.log_message(f"날짜 {date} 처리 완료")
            
            if self.is_running:
                self.progress_var.set(100)
                self.progress_text.config(text=f"완료! ({len(self.dates)}개 처리)")
                self.current_work_label.config(text="작업 완료")
                self.status_label.config(text="완료", foreground="green")
                self.log_message("모든 작업이 완료되었습니다.")
                
                messagebox.showinfo("완료", "모든 작업이 완료되었습니다.")
        
        except Exception as e:
            self.log_message(f"작업 중 오류 발생: {e}", "ERROR")
            messagebox.showerror("오류", f"작업 중 오류가 발생했습니다:\n{e}")
        finally:
            self.stop_work()
    
    def monitor_esc_key(self):
        while True:
            if keyboard.is_pressed('esc') and self.is_running:
                self.log_message("ESC 키가 감지되어 작업을 중단합니다.")
                self.stop_work()
                break
            time.sleep(0.1)
    
    def run(self):
        if tk is not None and hasattr(self, 'root'):
            self.root.mainloop()
        else:
            logging.info("TNLMacro.run() called in headless mode; GUI loop not started.")

if __name__ == "__main__":
    app = TNLMacro()
    app.run()