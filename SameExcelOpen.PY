import sys
import os
import subprocess
from tkinter import *
from tkinter import messagebox
# 드래그 앤 드롭 라이브러리 (pip install tkinterdnd2 필요)
from tkinterdnd2 import DND_FILES, TkinterDnD

class ExcelDualLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 이중 실행기 (CMD 방식)")
        self.root.geometry("600x450")
        self.root.resizable(False, False)

        # 파일 경로 저장 변수
        self.path1 = None
        self.path2 = None

        # --- UI 타이틀 ---
        Label(root, text="[배치파일 방식] 같은 이름의 엑셀 파일 2개 열기", 
              font=("맑은 고딕", 14, "bold"), fg="#333").pack(pady=15)

        # --- 메인 프레임 ---
        frame = Frame(root)
        frame.pack(expand=True, fill="both", padx=20, pady=5)

        # ================= [왼쪽 구역] =================
        self.frame1 = LabelFrame(frame, text=" [파일 1] ", font=("맑은 고딕", 11, "bold"), fg="blue")
        self.frame1.pack(side=LEFT, expand=True, fill="both", padx=5)
        
        # 리스트박스 (파일 정보 표시)
        self.lb1 = Listbox(self.frame1, bg="#f9f9f9", selectmode=SINGLE, font=("맑은 고딕", 10), height=5)
        self.lb1.pack(fill="both", padx=5, pady=5, expand=True)
        # 드래그 앤 드롭 이벤트 연결
        self.lb1.drop_target_register(DND_FILES)
        self.lb1.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, 1))

        # [X] 파일 제거 버튼
        self.btn_del1 = Button(self.frame1, text="❌ 파일 제거", command=self.clear_file1, 
                               bg="#ffCDD2", font=("맑은 고딕", 9))
        self.btn_del1.pack(fill="x", padx=5, pady=5)


        # ================= [오른쪽 구역] =================
        self.frame2 = LabelFrame(frame, text=" [파일 2] ", font=("맑은 고딕", 11, "bold"), fg="red")
        self.frame2.pack(side=RIGHT, expand=True, fill="both", padx=5)

        # 리스트박스
        self.lb2 = Listbox(self.frame2, bg="#f9f9f9", selectmode=SINGLE, font=("맑은 고딕", 10), height=5)
        self.lb2.pack(fill="both", padx=5, pady=5, expand=True)
        # 드래그 앤 드롭 이벤트 연결
        self.lb2.drop_target_register(DND_FILES)
        self.lb2.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, 2))

        # [X] 파일 제거 버튼
        self.btn_del2 = Button(self.frame2, text="❌ 파일 제거", command=self.clear_file2, 
                               bg="#ffCDD2", font=("맑은 고딕", 9))
        self.btn_del2.pack(fill="x", padx=5, pady=5)


        # ================= [실행 버튼] =================
        btn_frame = Frame(root)
        btn_frame.pack(fill="x", padx=20, pady=20)
        
        self.btn_run = Button(btn_frame, text="엑셀 2개 띄우기 (실행)", 
                              command=self.run_process, 
                              font=("맑은 고딕", 14, "bold"), bg="#28a745", fg="white", height=2)
        self.btn_run.pack(fill="x")

    # --- 기능 함수들 ---

    def clear_file1(self):
        """왼쪽 파일 삭제 기능"""
        self.path1 = None
        self.lb1.delete(0, END)
        self.frame1.config(bg="#f0f0f0") # 색상 초기화

    def clear_file2(self):
        """오른쪽 파일 삭제 기능"""
        self.path2 = None
        self.lb2.delete(0, END)
        self.frame2.config(bg="#f0f0f0") # 색상 초기화

    def clean_path(self, path):
        """드래그된 경로에서 {} 제거"""
        return path.strip('{}')

    def on_drop(self, event, area_num):
        """파일을 드래그했을 때 처리"""
        full_path = self.clean_path(event.data)
        
        if not os.path.isfile(full_path):
            messagebox.showerror("오류", "폴더가 아닌 '엑셀 파일'을 넣어주세요.")
            return

        filename = os.path.basename(full_path)
        
        if area_num == 1:
            self.path1 = os.path.abspath(full_path)
            self.lb1.delete(0, END)
            self.lb1.insert(0, f"📂 {filename}")
            self.lb1.insert(1, f"📍 {os.path.dirname(self.path1)}")
            self.frame1.config(bg="#E3F2FD") # 파란색 강조
        else:
            self.path2 = os.path.abspath(full_path)
            self.lb2.delete(0, END)
            self.lb2.insert(0, f"📂 {filename}")
            self.lb2.insert(1, f"📍 {os.path.dirname(self.path2)}")
            self.frame2.config(bg="#FFEBEE") # 빨간색 강조

    def launch_excel_cmd(self, file_path):
        """
        배치 파일(.bat)의 방식을 그대로 차용
        명령어: start excel /e /x "경로"
        /x : 새로운 엑셀 인스턴스(프로세스) 생성 -> 동시 실행 가능하게 함
        /e : 시작 화면 없이 열기
        """
        try:
            # subprocess를 이용해 윈도우 명령어로 실행 (스레드 오류 없음)
            # shell=True를 쓰면 'start' 명령어를 cmd에서 치는 것과 같음
            subprocess.Popen(f'start excel /e /x "{file_path}"', shell=True)
        except Exception as e:
            messagebox.showerror("실행 실패", f"엑셀을 실행하지 못했습니다.\n{e}")

    def run_process(self):
        if not self.path1 or not self.path2:
            messagebox.showwarning("경고", "파일 2개를 모두 지정해주세요.")
            return

        if self.path1 == self.path2:
            messagebox.showwarning("주의", "경로가 완전히 같은 파일입니다.\n서로 다른 위치의 파일을 사용해주세요.")
            return

        # 배치 파일 방식 실행 (스레드 필요 없음, 즉시 실행)
        self.launch_excel_cmd(self.path1)
        self.launch_excel_cmd(self.path2)

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ExcelDualLauncher(root)
    root.mainloop()