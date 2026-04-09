import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog, simpledialog
import os
import subprocess
import sys
import json
import socket
import ctypes
import shutil

# 드래그 앤 드롭 라이브러리 체크
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    print("tkinterdnd2 라이브러리가 필요합니다. 'pip install tkinterdnd2'를 실행하세요.")
    sys.exit(1)

# --- 경로 설정 ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 콘솔 숨김 (Windows)
if sys.platform == "win32":
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd != 0:
            ctypes.windll.user32.ShowWindow(hwnd, 0) 
    except Exception:
        pass

# 레거시 및 설정 파일 경로
LEGACY_SOURCE_LIST_FILE = os.path.join(BASE_DIR, "COPY_dir_list.txt")
LEGACY_DEST_LIST_FILE = os.path.join(BASE_DIR, "PASTE_dir_list.txt")
LEGACY_SERVER_PATH_LIST_FILE = os.path.join(BASE_DIR, "SERVER_path_list.txt")
TABS_CONFIG_FILE = os.path.join(BASE_DIR, "tabs_config.json") 

TARGET_PROCESSES = [
    "LoginServer.exe", "GameServer.exe",
    "LoginServer_TW.exe", "GameServer_TW.exe"
]

# =============================================================================
# [통합 탭 클래스]
# =============================================================================
class AllInOneTab(tk.Frame):
    def __init__(self, parent, app_ref, initial_state=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.app_ref = app_ref 
        self.initial_state = initial_state or {} 
        
        self.font_large = ("맑은 고딕", 11, "bold")
        self.font_normal = ("맑은 고딕", 10)
        self.font_small = ("맑은 고딕", 9)

        # ---------------------------------------------------------------------
        # [데이터 초기화]
        # ---------------------------------------------------------------------
        self.src_items = self.initial_state.get("src_list", [])
        self.dst_items = self.initial_state.get("dst_list", [])
        self.server_items = self.initial_state.get("server_list", [])
        
        # 섹션 4 기본값
        default_files = ["Localization.xlsm", "LocalizationName.xlsm", "IntroLocalization.xlsm"]
        self.specific_items = self.initial_state.get("specific_list", default_files)

        # 마이그레이션 로직
        if not self.src_items and os.path.exists(LEGACY_SOURCE_LIST_FILE):
             with open(LEGACY_SOURCE_LIST_FILE, "r", encoding="utf-8") as f:
                 self.src_items = [l.strip() for l in f if l.strip()]
        
        if not self.dst_items and os.path.exists(LEGACY_DEST_LIST_FILE):
             with open(LEGACY_DEST_LIST_FILE, "r", encoding="utf-8") as f:
                 self.dst_items = [l.strip() for l in f if l.strip()]

        if not self.server_items and os.path.exists(LEGACY_SERVER_PATH_LIST_FILE):
             with open(LEGACY_SERVER_PATH_LIST_FILE, "r", encoding="utf-8") as f:
                 self.server_items = [l.strip() for l in f if l.strip()]
             if BASE_DIR not in self.server_items:
                 self.server_items.insert(0, BASE_DIR)
        elif not self.server_items:
            self.server_items = [BASE_DIR]

        # UI 생성
        self._create_copy_ui()   
        self._create_server_ui()
        self._create_ip_ui() 
        self._create_specific_copy_ui()
        
        # 데이터 로드
        self.refresh_data()
        self._apply_saved_selection()

    def refresh_data(self):
        """UI 리프레시"""
        self._update_copy_ui_from_memory()
        self._update_server_ui_from_memory()
        self._update_specific_ui_from_memory()

    def get_current_state(self):
        return {
            "src": self.combo_src.get(),
            "dst": self.combo_dst.get(),
            "server": self.combo_server.get(),
            "ip": self.entry_ip.get(),
            "spec_src": self.combo_spec_src.get(),
            "spec_dst": self.combo_spec_dst.get(),
            
            "src_list": self.src_items,
            "dst_list": self.dst_items,
            "server_list": self.server_items,
            "specific_list": self.specific_items
        }

    def _apply_saved_selection(self):
        if not self.initial_state: 
            self._fill_local_ip()
            return
        
        src = self.initial_state.get("src", "")
        dst = self.initial_state.get("dst", "")
        srv = self.initial_state.get("server", "")
        ip = self.initial_state.get("ip", "")
        spec_src = self.initial_state.get("spec_src", "")
        spec_dst = self.initial_state.get("spec_dst", "")

        self._ensure_combo_value(self.combo_src, src)
        self._ensure_combo_value(self.combo_dst, dst)
        self._ensure_combo_value(self.combo_server, srv)
        self._ensure_combo_value(self.combo_spec_src, spec_src)
        self._ensure_combo_value(self.combo_spec_dst, spec_dst)
        
        if ip:
            self.entry_ip.delete(0, tk.END)
            self.entry_ip.insert(0, ip)
        else:
            self._fill_local_ip()

    def _ensure_combo_value(self, combo, value):
        if not value: return
        curr_vals = list(combo['values'])
        if value not in curr_vals:
            curr_vals.append(value)
            combo['values'] = curr_vals
        combo.set(value)

    def _on_interaction(self, event=None):
        if self.app_ref:
            self.app_ref.save_tabs_state()

    # -------------------------------------------------------------------------
    # 1. 파일 복사 UI
    # -------------------------------------------------------------------------
    def _create_copy_ui(self):
        copy_frame = tk.LabelFrame(self, text="[1. 파일 복사 관리(한국 :TestServer / 대만 TestServer_TW)]", font=self.font_large, padx=5, pady=5)
        copy_frame.pack(pady=5, padx=10, fill="x")

        list_frame = tk.Frame(copy_frame)
        list_frame.pack(fill="x", pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.columnconfigure(2, weight=1)

        # Source
        tk.Label(list_frame, text="복사할 폴더 목록:", font=self.font_small, fg="blue").grid(row=0, column=0, sticky="w")
        src_scroll_frame = tk.Frame(list_frame)
        src_scroll_frame.grid(row=1, column=0, sticky="nsew", padx=2)
        scrollbar_src = tk.Scrollbar(src_scroll_frame)
        scrollbar_src.pack(side="right", fill="y")
        self.list_src = tk.Listbox(src_scroll_frame, height=3, font=self.font_small, selectmode=tk.SINGLE, yscrollcommand=scrollbar_src.set)
        self.list_src.pack(side="left", fill="both", expand=True)
        scrollbar_src.config(command=self.list_src.yview)
        
        btn_src = tk.Frame(list_frame)
        btn_src.grid(row=1, column=1, padx=2)
        tk.Button(btn_src, text="추가", command=lambda: self._add_copy_path("source"), width=6, bg="#E8F5E9").pack(pady=1)
        tk.Button(btn_src, text="삭제", command=lambda: self._del_copy_path("source"), width=6, bg="#FFEBEE").pack(pady=1)

        # Dest
        tk.Label(list_frame, text="붙여넣을 폴더 목록:", font=self.font_small, fg="red").grid(row=0, column=2, sticky="w", padx=(10,0))
        dst_scroll_frame = tk.Frame(list_frame)
        dst_scroll_frame.grid(row=1, column=2, sticky="nsew", padx=2)
        scrollbar_dst = tk.Scrollbar(dst_scroll_frame)
        scrollbar_dst.pack(side="right", fill="y")
        self.list_dst = tk.Listbox(dst_scroll_frame, height=3, font=self.font_small, selectmode=tk.SINGLE, yscrollcommand=scrollbar_dst.set)
        self.list_dst.pack(side="left", fill="both", expand=True)
        scrollbar_dst.config(command=self.list_dst.yview)

        btn_dst = tk.Frame(list_frame)
        btn_dst.grid(row=1, column=3, padx=2)
        tk.Button(btn_dst, text="추가", command=lambda: self._add_copy_path("dest"), width=6, bg="#E8F5E9").pack(pady=1)
        tk.Button(btn_dst, text="삭제", command=lambda: self._del_copy_path("dest"), width=6, bg="#FFEBEE").pack(pady=1)

        # Exec
        exec_frame = tk.Frame(copy_frame)
        exec_frame.pack(fill="x", pady=5)
        exec_frame.columnconfigure(1, weight=1)

        tk.Label(exec_frame, text="복사:", font=self.font_normal, fg="blue").grid(row=0, column=0, sticky="e")
        self.combo_src = ttk.Combobox(exec_frame, font=self.font_normal, state="readonly")
        self.combo_src.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.combo_src.bind("<<ComboboxSelected>>", self._on_interaction)

        tk.Label(exec_frame, text="붙여넣기:", font=self.font_normal, fg="red").grid(row=1, column=0, sticky="e")
        self.combo_dst = ttk.Combobox(exec_frame, font=self.font_normal, state="readonly")
        self.combo_dst.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.combo_dst.bind("<<ComboboxSelected>>", self._on_interaction)

        tk.Button(exec_frame, text="▶ 전체 복사", command=self._run_copy, 
                  bg="#e3f2fd", width=12, height=2, font=("맑은 고딕", 10, "bold")).grid(row=0, column=2, rowspan=2, padx=10)

    def _update_copy_ui_from_memory(self):
        self.list_src.delete(0, tk.END)
        for item in self.src_items: self.list_src.insert(tk.END, item)
        self.combo_src['values'] = self.src_items
        
        self.list_dst.delete(0, tk.END)
        for item in self.dst_items: self.list_dst.insert(tk.END, item)
        self.combo_dst['values'] = self.dst_items
        
        if hasattr(self, 'combo_spec_src'): self.combo_spec_src['values'] = self.src_items
        if hasattr(self, 'combo_spec_dst'): self.combo_spec_dst['values'] = self.dst_items

    def _add_copy_path(self, type_):
        path = filedialog.askdirectory(title="폴더 선택")
        if not path: return
        target_list = self.src_items if type_ == "source" else self.dst_items
        if path not in target_list:
            target_list.append(path)
            self._update_copy_ui_from_memory()
            self._on_interaction()

    def _del_copy_path(self, type_):
        target_listbox = self.list_src if type_ == "source" else self.list_dst
        target_list = self.src_items if type_ == "source" else self.dst_items
        idx = target_listbox.curselection()
        if not idx: return
        val = target_listbox.get(idx)
        if val in target_list:
            target_list.remove(val)
            self._update_copy_ui_from_memory()
            self._on_interaction()

    def _run_copy(self):
        src = self.combo_src.get()
        dst = self.combo_dst.get()
        if not src or not dst: return messagebox.showwarning("경고", "경로를 선택하세요.")
        try:
            cmd = f'robocopy "{src}" "{dst}" /E /MT:8 /R:1 /W:1 /NFL /NDL'
            proc = subprocess.run(cmd, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if proc.returncode <= 7: messagebox.showinfo("완료", "전체 복사 성공!")
            else: messagebox.showerror("실패", f"오류 코드: {proc.returncode}")
        except Exception as e: messagebox.showerror("에러", str(e))

    # -------------------------------------------------------------------------
    # 2. 서버 제어 UI
    # -------------------------------------------------------------------------
    def _create_server_ui(self):
        server_frame = tk.LabelFrame(self, text="[2. 서버 제어(한국 :TestServer / 대만 TestServer_TW)]", font=self.font_large, padx=5, pady=5)
        server_frame.pack(pady=5, padx=10, fill="x")

        path_line = tk.Frame(server_frame)
        path_line.pack(fill="x", pady=2)
        
        tk.Label(path_line, text="서버 위치:", font=self.font_normal).pack(side="left")
        self.combo_server = ttk.Combobox(path_line, font=self.font_normal, state="readonly")
        self.combo_server.pack(side="left", padx=5, fill="x", expand=True) 
        self.combo_server.bind("<<ComboboxSelected>>", self._on_interaction)
        
        tk.Button(path_line, text="추가", command=self._add_server_path, bg="#E8F5E9").pack(side="left", padx=1)
        tk.Button(path_line, text="삭제", command=self._del_server_path, bg="#FFEBEE").pack(side="left", padx=1)

        ctrl_line = tk.Frame(server_frame)
        ctrl_line.pack(fill="x", pady=5)
        
        tk.Button(ctrl_line, text="시작 / 재시작", bg="lightgreen", command=self._start_server, font=("맑은 고딕", 10, "bold")).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(ctrl_line, text="종료", bg="salmon", command=self._stop_server, font=("맑은 고딕", 10, "bold")).pack(side="left", fill="x", expand=True, padx=5)

    def _update_server_ui_from_memory(self):
        self.combo_server['values'] = self.server_items
        current = self.combo_server.get()
        if current and current not in self.server_items: self.combo_server.set('')
        if not self.combo_server.get() and self.server_items: self.combo_server.current(0)

    def _add_server_path(self):
        path = filedialog.askdirectory(title="서버 폴더 선택")
        if not path: return
        if path not in self.server_items:
            self.server_items.append(path)
            self._update_server_ui_from_memory()
            self.combo_server.set(path)
            self._on_interaction()

    def _del_server_path(self):
        curr = self.combo_server.get()
        if not curr: return
        if curr == BASE_DIR: return messagebox.showwarning("불가", "기본 경로는 삭제 불가")
        if curr in self.server_items:
            self.server_items.remove(curr)
            self.combo_server.set('')
            self._update_server_ui_from_memory()
            self._on_interaction()

    def _start_server(self):
        self._stop_server()
        cwd = self.combo_server.get()
        if not cwd: return messagebox.showwarning("오류", "서버 경로를 선택하세요.")
        login_exe = "LoginServer.exe"
        game_exe = "GameServer.exe"
        if "_TW" in cwd:
            login_exe = "LoginServer_TW.exe"
            game_exe = "GameServer_TW.exe"
        try:
            subprocess.run(f'start "" "{login_exe}" 9011', shell=True, cwd=cwd, creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run(f'start "" "{game_exe}" 9101', shell=True, cwd=cwd, creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception as e: messagebox.showerror("오류", str(e))

    def _stop_server(self):
        for proc in TARGET_PROCESSES:
            subprocess.run(f'taskkill /F /IM "{proc}"', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)

    # -------------------------------------------------------------------------
    # 3. IP 주소 UI
    # -------------------------------------------------------------------------
    def _create_ip_ui(self):
        ip_frame = tk.LabelFrame(self, text="[3. IP 주소 변경]", font=self.font_large, padx=5, pady=5)
        ip_frame.pack(pady=5, padx=10, fill="x")

        input_frame = tk.Frame(ip_frame)
        input_frame.pack(fill="x", pady=2)

        tk.Label(input_frame, text="IP:", font=("맑은 고딕", 10)).pack(side="left", padx=5)
        self.entry_ip = tk.Entry(input_frame, font=("맑은 고딕", 10))
        self.entry_ip.pack(side="left", padx=5, fill="x", expand=True)
        self.entry_ip.bind("<FocusOut>", self._on_interaction)

        tk.Button(input_frame, text="내 IP", command=self._fill_local_ip, bg="#e0f7fa", font=self.font_small).pack(side="left", padx=5)

        btn_frame = tk.Frame(ip_frame)
        btn_frame.pack(fill="x", pady=5)
        tk.Button(btn_frame, text="127.0.0.1 ➔ IP", bg="#fff9c4", command=lambda: self._change_ip(True)).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(btn_frame, text="IP ➔ 127.0.0.1", bg="#ffe0b2", command=lambda: self._change_ip(False)).pack(side="left", fill="x", expand=True, padx=2)

    def _fill_local_ip(self):
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(('10.255.255.255', 1))
            IP = s.getsockname()[0]
        except: IP = '127.0.0.1'
        finally: s.close()
        self.entry_ip.delete(0, tk.END)
        self.entry_ip.insert(0, IP)
        self._on_interaction()

    def _change_ip(self, to_target):
        server_path = self.combo_server.get()
        source_path = self.combo_src.get()
        target_ip = self.entry_ip.get().strip()
        if not target_ip: return messagebox.showwarning("오류", "IP 입력 필요")
        if not server_path or not source_path: return messagebox.showwarning("오류", "경로 선택 필요")

        src_t = "127.0.0.1" if to_target else target_ip
        dst_t = target_ip if to_target else "127.0.0.1"

        logs = []
        for f in ["9011_LoginServer_Config.json", "9101_GameServer_Config.json"]:
            p = os.path.join(server_path, "Config", f)
            logs.append(f"[서버] {f}: {self._replace_file(p, src_t, dst_t)}")

        norm_src = os.path.normpath(source_path)
        base = os.path.dirname(norm_src) if os.path.basename(norm_src).lower() == "table" else norm_src.rsplit("Table", 1)[0] if "Table" in norm_src else norm_src
        cli_p = os.path.join(base, "Assets", "AddressableResources", "Scriptable", "LoginServerData.asset")
        
        res = self._replace_file(cli_p, src_t, dst_t)
        logs.append(f"[클라] Asset: {res}")
        messagebox.showinfo("결과", "\n".join(logs))

    def _replace_file(self, path, old, new):
        if not os.path.exists(path): return "파일 없음"
        try:
            enc = 'utf-8'
            try: 
                with open(path, 'r', encoding='utf-8') as f: c = f.read()
            except: 
                enc = 'cp949'
                with open(path, 'r', encoding='cp949') as f: c = f.read()
            
            if old not in c: return "이미 변경됨/없음" if new in c else "대상 텍스트 없음"
            with open(path, 'w', encoding=enc) as f: f.write(c.replace(old, new))
            return "성공"
        except Exception as e: return f"에러: {e}"

    # -------------------------------------------------------------------------
    # 4. 지정 파일/폴더 복사 UI (드래그 앤 드롭 지원)
    # -------------------------------------------------------------------------
    def _create_specific_copy_ui(self):
        spec_frame = tk.LabelFrame(self, text="[4. 번역 때문에 제작 지정 파일/폴더 복사]", font=self.font_large, padx=5, pady=5)
        spec_frame.pack(pady=5, padx=10, fill="both", expand=True)

        # 경로 선택 영역
        path_frame = tk.Frame(spec_frame)
        path_frame.pack(fill="x", pady=2)
        path_frame.columnconfigure(1, weight=1)

        tk.Label(path_frame, text="복사:", fg="blue", font=self.font_small).grid(row=0, column=0, sticky="e")
        self.combo_spec_src = ttk.Combobox(path_frame, font=self.font_small, state="readonly")
        self.combo_spec_src.grid(row=0, column=1, sticky="ew", padx=5)
        self.combo_spec_src.bind("<<ComboboxSelected>>", self._on_interaction)

        tk.Label(path_frame, text="붙여넣기:", fg="red", font=self.font_small).grid(row=1, column=0, sticky="e")
        self.combo_spec_dst = ttk.Combobox(path_frame, font=self.font_small, state="readonly")
        self.combo_spec_dst.grid(row=1, column=1, sticky="ew", padx=5)
        self.combo_spec_dst.bind("<<ComboboxSelected>>", self._on_interaction)

        # 파일 리스트 영역
        list_area = tk.Frame(spec_frame)
        list_area.pack(fill="both", expand=True, pady=5)
        
        tk.Label(list_area, text="▼ 아래 리스트에 파일을 드래그&드롭하여 추가 가능", font=self.font_small, fg="gray").pack(anchor="w")
        
        scroll_f = tk.Frame(list_area)
        scroll_f.pack(fill="both", expand=True)
        sb = tk.Scrollbar(scroll_f)
        sb.pack(side="right", fill="y")
        self.list_specific = tk.Listbox(scroll_f, height=3, font=self.font_small, yscrollcommand=sb.set)
        self.list_specific.pack(side="left", fill="both", expand=True)
        sb.config(command=self.list_specific.yview)

        # 드래그 앤 드롭 등록
        self.list_specific.drop_target_register(DND_FILES)
        self.list_specific.dnd_bind('<<Drop>>', self._on_drop_files)

        # 버튼 영역
        btn_area = tk.Frame(spec_frame)
        btn_area.pack(fill="x", pady=5)
        
        # [변경] "파일 선택 추가" 버튼도 삭제함 (드래그앤드롭 사용)
        tk.Button(btn_area, text="선택 삭제", command=self._del_spec_item, bg="#FFEBEE").pack(side="left", padx=2)
        
        tk.Button(btn_area, text="▶ 지정 복사 실행", command=self._run_specific_copy, 
                  bg="#e3f2fd", font=("맑은 고딕", 10, "bold")).pack(side="right", padx=5)

    def _on_drop_files(self, event):
        """드래그 앤 드롭 이벤트 핸들러"""
        if not event.data: return
        
        # 경로 파싱
        file_paths = self.master.tk.splitlist(event.data)
        
        added_count = 0
        for p in file_paths:
            name = os.path.basename(p)
            if name and name not in self.specific_items:
                self.specific_items.append(name)
                added_count += 1
        
        if added_count > 0:
            self._update_specific_ui_from_memory()
            self._on_interaction()

    def _update_specific_ui_from_memory(self):
        self.combo_spec_src['values'] = self.src_items
        self.combo_spec_dst['values'] = self.dst_items
        
        self.list_specific.delete(0, tk.END)
        for item in self.specific_items:
            self.list_specific.insert(tk.END, item)

    def _del_spec_item(self):
        idx = self.list_specific.curselection()
        if not idx: return
        val = self.list_specific.get(idx)
        if val in self.specific_items:
            self.specific_items.remove(val)
            self._update_specific_ui_from_memory()
            self._on_interaction()

    def _run_specific_copy(self):
        src_root = self.combo_spec_src.get()
        dst_root = self.combo_spec_dst.get()
        
        if not src_root or not dst_root:
            return messagebox.showwarning("경고", "Source와 Dest 경로를 선택해주세요.")
            
        count = 0
        errors = []
        
        for item in self.specific_items:
            s_path = os.path.join(src_root, item)
            d_path = os.path.join(dst_root, item)
            
            try:
                if os.path.isdir(s_path):
                    shutil.copytree(s_path, d_path, dirs_exist_ok=True)
                    count += 1
                elif os.path.isfile(s_path):
                    os.makedirs(os.path.dirname(d_path), exist_ok=True)
                    shutil.copy2(s_path, d_path)
                    count += 1
                else:
                    errors.append(f"없음: {item}")
            except Exception as e:
                errors.append(f"에러({item}): {str(e)}")
        
        msg = f"총 {count}개 항목 복사 완료."
        if errors:
            msg += "\n\n[미처리 항목]\n" + "\n".join(errors)
            messagebox.showwarning("완료 (일부 실패)", msg)
        else:
            messagebox.showinfo("성공", msg)


# =============================================================================
# [메인 앱]
# =============================================================================
class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("개인 서버 도우미")
        self.root.geometry("1080x750") 
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        style = ttk.Style()
        style.theme_use('default')
        style.configure('TNotebook.Tab', padding=[20, 10], font=('맑은 고딕', 11, 'bold'))

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both", padx=5, pady=5)
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        self.context_menu = tk.Menu(root, tearoff=0)
        self.context_menu.add_command(label="탭 이름 변경", command=self.rename_tab)
        self.context_menu.add_command(label="탭 삭제", command=self.delete_tab)
        self.notebook.bind("<Button-3>", self.show_context_menu)

        self.tabs_list = []
        self.load_tabs()

    def show_context_menu(self, event):
        try:
            index = self.notebook.index(f"@{event.x},{event.y}")
            if index == len(self.notebook.tabs()) - 1:
                return
            self.notebook.select(index)
            self.context_menu.post(event.x_root, event.y_root)
        except: pass

    def on_tab_changed(self, event):
        try:
            total_tabs = self.notebook.tabs()
            if not total_tabs: return
            selected_idx = self.notebook.index(self.notebook.select())
            last_idx = len(total_tabs) - 1
            if selected_idx == last_idx and self.notebook.tab(last_idx, "text") == "+":
                self.add_tab("새 작업")
        except Exception as e:
            print(e)

    def add_tab(self, name, initial_state=None):
        tabs = self.notebook.tabs()
        has_plus_tab = False
        if tabs:
            last_idx = len(tabs) - 1
            if self.notebook.tab(last_idx, "text") == "+":
                has_plus_tab = True

        tab = AllInOneTab(self.notebook, app_ref=self, initial_state=initial_state) 
        
        if has_plus_tab:
            self.notebook.insert(len(tabs)-1, tab, text=name)
        else:
            self.notebook.add(tab, text=name)

        self.notebook.select(tab)
        self.tabs_list.append(tab)
        self.save_tabs_state()

    def rename_tab(self):
        idx = self.notebook.select()
        if self.notebook.tab(idx, "text") == "+": return
        new_name = simpledialog.askstring("이름 변경", "새 탭 이름:", initialvalue=self.notebook.tab(idx, "text"))
        if new_name:
            self.notebook.tab(idx, text=new_name)
            self.save_tabs_state()

    def delete_tab(self):
        curr_widget_name = self.notebook.select()
        idx = self.notebook.index(curr_widget_name)
        if self.notebook.tab(idx, "text") == "+": return
        if len(self.tabs_list) <= 1: 
            return messagebox.showwarning("경고", "최소 하나의 작업 탭은 있어야 합니다.")
        
        if messagebox.askyesno("삭제", "현재 탭을 삭제하시겠습니까?"):
            curr_widget = self.notebook.nametowidget(curr_widget_name)
            if curr_widget in self.tabs_list: 
                self.tabs_list.remove(curr_widget)
            self.notebook.forget(idx)
            self.save_tabs_state()

    def save_tabs_state(self):
        data = []
        for tab in self.tabs_list:
            try:
                idx = self.notebook.index(tab)
                data.append({
                    "name": self.notebook.tab(idx, "text"),
                    "state": tab.get_current_state()
                })
            except: continue
        
        curr_idx = 0
        try: 
            curr_widget = self.notebook.select()
            curr_idx = self.notebook.index(curr_widget)
            if self.notebook.tab(curr_idx, "text") == "+":
                curr_idx = max(0, curr_idx - 1)
        except: pass

        with open(TABS_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_selected_index": curr_idx, "tabs": data}, f, ensure_ascii=False, indent=2)

    def load_tabs(self):
        raw_data = {"tabs": [], "last_selected_index": 0}
        if os.path.exists(TABS_CONFIG_FILE):
            try:
                with open(TABS_CONFIG_FILE, "r", encoding="utf-8") as f:
                    content = json.load(f)
                    if isinstance(content, list): raw_data["tabs"] = content
                    elif isinstance(content, dict): raw_data = content
            except: pass
        
        if raw_data.get("tabs"):
            for item in raw_data["tabs"]:
                if isinstance(item, dict): self.add_tab(item.get("name", "탭"), item.get("state"))
                else: self.add_tab(str(item)) 
        else:
            self.add_tab("Default")

        dummy_frame = tk.Frame(self.notebook)
        self.notebook.add(dummy_frame, text="+")

        try: 
            target_idx = raw_data.get("last_selected_index", 0)
            if target_idx >= len(self.tabs_list):
                target_idx = len(self.tabs_list) - 1
            self.notebook.select(self.tabs_list[target_idx])
        except: pass
            
    def on_closing(self):
        self.save_tabs_state()
        self.root.destroy()

if __name__ == "__main__":
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
        messagebox.showerror("오류", "tkinterdnd2 모듈이 없습니다.\n'pip install tkinterdnd2'를 설치해주세요.")
    
    app = MainApp(root)
    root.mainloop()