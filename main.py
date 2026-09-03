import os
import tempfile
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import pyzipper

# 외관 및 테마 설정 (Apple / Toss 미니멀 라이트 모드)
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

class ReportMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("보고서 데이터 통합 도구")
        self.geometry("620 x 760")
        self.resizable(False, False)
        
        # Toss / Apple 시그니처 연회색 배경
        self.configure(fg_color="#F2F4F6")

        self.selected_files = []
        
        # 기본 저장 경로: 사용자 바탕화면
        self.save_dir_path = str(Path.home() / "Desktop")
        if not os.path.exists(self.save_dir_path):
            self.save_dir_path = str(Path.home())

        self._init_ui()

    def _init_ui(self):
        # 1. 상단 타이틀 카드 (Toss 볼드 헤드라인)
        self.header_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=18)
        self.header_card.pack(fill="x", padx=20, pady=(18, 10))

        self.title_label = ctk.CTkLabel(
            self.header_card,
            text="보고서 데이터 통합",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=22, weight="bold"),
            text_color="#191F28"
        )
        self.title_label.pack(anchor="w", padx=24, pady=(18, 4))

        self.desc_label = ctk.CTkLabel(
            self.header_card,
            text="압축 파일(.zip)과 일반 엑셀/CSV 문서를 컬럼 기준으로 자동 병합합니다.",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13),
            text_color="#8B95A1"
        )
        self.desc_label.pack(anchor="w", padx=24, pady=(0, 18))

        # 2. 메인 설정 카드
        self.content_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=18)
        self.content_card.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        # [섹션 1] 파일 선택 버튼 영역
        self.btn_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=24, pady=(18, 8))

        self.file_btn = ctk.CTkButton(
            self.btn_frame,
            text="+ 취합할 파일 추가 (.zip, .xlsx, .csv)",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            fg_color="#3182F6",      # Toss Blue
            hover_color="#1B64DA",
            corner_radius=10,
            height=38,
            command=self.select_files
        )
        self.file_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.clear_btn = ctk.CTkButton(
            self.btn_frame,
            text="목록 비우기",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#4E5968",
            corner_radius=10,
            width=85,
            height=38,
            command=self.clear_file_list
        )
        self.clear_btn.pack(side="right")

        # 파일 목록 박스
        self.file_listbox = tk.Listbox(
            self.content_card,
            height=5,
            bd=0,
            bg="#F9FAFB",
            fg="#191F28",
            highlightthickness=1,
            highlightcolor="#3182F6",
            highlightbackground="#E5E8EB",
            selectbackground="#E8F3FF",
            selectforeground="#1B64DA",
            font=("Apple SD Gothic Neo", 10)
        )
        self.file_listbox.pack(fill="both", expand=True, padx=24, pady=(0, 14))

        # [섹션 2] 저장할 파일명
        self.name_label = ctk.CTkLabel(
            self.content_card,
            text="저장할 엑셀 파일 이름",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            text_color="#333D4B"
        )
        self.name_label.pack(anchor="w", padx=24, pady=(0, 5))

        self.filename_entry = ctk.CTkEntry(
            self.content_card,
            placeholder_text="통합_보고서.xlsx",
            corner_radius=10,
            height=38,
            fg_color="#F9FAFB",
            border_color="#E5E8EB",
            border_width=1,
            text_color="#191F28"
        )
        self.filename_entry.insert(0, "통합_보고서.xlsx")
        self.filename_entry.pack(fill="x", padx=24, pady=(0, 14))

        # [섹션 3] 저장 경로
        self.path_label = ctk.CTkLabel(
            self.content_card,
            text="저장 위치 (기본값: 바탕화면)",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            text_color="#333D4B"
        )
        self.path_label.pack(anchor="w", padx=24, pady=(0, 5))

        self.path_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.path_frame.pack(fill="x", padx=24, pady=(0, 14))

        self.path_entry = ctk.CTkEntry(
            self.path_frame,
            corner_radius=10,
            height=38,
            fg_color="#F9FAFB",
            border_color="#E5E8EB",
            border_width=1,
            text_color="#191F28"
        )
        self.path_entry.insert(0, self.save_dir_path)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.path_btn = ctk.CTkButton(
            self.path_frame,
            text="폴더 변경",
            width=80,
            height=38,
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=10,
            command=self.change_directory
        )
        self.path_btn.pack(side="right")

        # [섹션 4] 진행 상황 및 프로그레스 바
        self.progress_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=24, pady=(0, 16))

        self.progress_status_label = ctk.CTkLabel(
            self.progress_frame,
            text="대기 중",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12),
            text_color="#8B95A1"
        )
        self.progress_status_label.pack(side="left")

        self.progress_percent_label = ctk.CTkLabel(
            self.progress_frame,
            text="0%",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12, weight="bold"),
            text_color="#3182F6"
        )
        self.progress_percent_label.pack(side="right")

        self.progress_bar = ctk.CTkProgressBar(
            self.content_card,
            corner_radius=6,
            height=8,
            progress_color="#3182F6",
            fg_color="#E5E8EB"
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=24, pady=(0, 18))

        # 3. 하단 액션 버튼
        self.run_btn = ctk.CTkButton(
            self,
            text="데이터 취합 시작하기",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=15, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            corner_radius=14,
            height=48,
            command=self.start_merge_thread
        )
        self.run_btn.pack(fill="x", padx=20, pady=(0, 10))

        # 4. 맨 하단 라이선스 및 면책 문구
        self.footer_label = ctk.CTkLabel(
            self,
            text="본 프로그램은 업무 지원을 위해 제작된 프로그램으로, MIT 라이선스를 따릅니다.\n(수정, 배포 자유로우며 모든 책임은 사용자 본인에게 귀속됩니다.)",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=10),
            text_color="#8B95A1",
            justify="center"
        )
        self.footer_label.pack(padx=20, pady=(0, 14))

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="취합할 파일들을 선택하세요 (복수 선택 가능)",
            filetypes=[
                ("모든 지원 파일", "*.zip *.xlsx *.xls *.csv"),
                ("압축 파일 (.zip)", "*.zip"),
                ("엑셀 파일 (.xlsx, .xls)", "*.xlsx *.xls"),
                ("CSV 파일 (.csv)", "*.csv"),
                ("모든 파일", "*.*")
            ]
        )
        if files:
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
            self._refresh_listbox()

    def clear_file_list(self):
        self.selected_files.clear()
        self._refresh_listbox()
        self._set_progress(0, "대기 중")

    def _refresh_listbox(self):
        self.file_listbox.delete(0, tk.END)
        for f in self.selected_files:
            ext = os.path.splitext(f)[1].lower()
            icon = "📦" if ext == ".zip" else "📄"
            self.file_listbox.insert(tk.END, f" {icon}  {os.path.basename(f)}")

    def change_directory(self):
        folder = filedialog.askdirectory(initialdir=self.path_entry.get(), title="저장할 폴더 선택")
        if folder:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder)

    def ask_zip_password(self, filename):
        dialog = ctk.CTkInputDialog(
            text=f"[{os.path.basename(filename)}]\n비밀번호가 잠겨있습니다.\n압축 해제를 위한 번호를 입력해 주세요.",
            title="비밀번호 필요"
        )
        return dialog.get_input()

    def _set_progress(self, percent, text):
        def update():
            clamped = max(0, min(100, percent))
            self.progress_bar.set(clamped / 100.0)
            self.progress_percent_label.configure(text=f"{clamped}%")
            self.progress_status_label.configure(text=text)
        self.after(0, update)

    def start_merge_thread(self):
        if not self.selected_files:
            messagebox.showwarning("안내", "병합할 파일(.zip 또는 엑셀 문서)을 먼저 추가해 주세요.")
            return

        out_name = self.filename_entry.get().strip()
        if not out_name:
            out_name = "통합_보고서.xlsx"
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"

        save_dir = self.path_entry.get().strip()
        if not os.path.exists(save_dir):
            messagebox.showerror("오류", "유효한 저장 경로를 지정해 주세요.")
            return

        out_path = os.path.join(save_dir, out_name)
        self.run_btn.configure(state="disabled", text="데이터를 취합하는 중입니다...")
        self._set_progress(0, "취합 준비 중...")

        threading.Thread(target=self.process_merge, args=(out_path,), daemon=True).start()

    def process_merge(self, out_path):
        dataframes = []
        excel_files_to_read = []

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_files = [f for f in self.selected_files if os.path.splitext(f)[1].lower() == ".zip"]
                direct_files = [f for f in self.selected_files if os.path.splitext(f)[1].lower() in [".xlsx", ".xls", ".csv"]]

                # 1단계: 압축 파일 해제 (0% ~ 30%)
                for idx, zip_path in enumerate(zip_files):
                    current_percent = int((idx / max(1, len(zip_files))) * 30)
                    self._set_progress(current_percent, f"압축 해제 중 ({idx + 1}/{len(zip_files)})...")

                    sub_dir = os.path.join(temp_dir, f"zip_extracted_{idx}")
                    os.makedirs(sub_dir, exist_ok=True)

                    extracted = False
                    password = None

                    while not extracted:
                        try:
                            with pyzipper.AESZipFile(zip_path) as zf:
                                if password:
                                    zf.setpassword(password.encode("utf-8"))
                                zf.extractall(path=sub_dir)
                            extracted = True
                        except (RuntimeError, pyzipper.BadZipFile):
                            password = self.ask_zip_password(zip_path)
                            if not password:
                                self.reset_ui()
                                messagebox.showinfo("취소", f"[{os.path.basename(zip_path)}] 암호 입력을 취소하여 작업을 중단합니다.")
                                return

                    for root, _, files in os.walk(sub_dir):
                        for file in files:
                            f_ext = os.path.splitext(file)[1].lower()
                            if f_ext in [".xlsx", ".xls", ".csv"]:
                                excel_files_to_read.append(os.path.join(root, file))

                # 일반 파일 추가
                excel_files_to_read.extend(direct_files)

                total_docs = len(excel_files_to_read)
                if total_docs == 0:
                    self.reset_ui()
                    messagebox.showwarning("경고", "취합할 수 있는 유효한 엑셀 또는 CSV 데이터가 없습니다.")
                    return

                # 2단계: 데이터 읽기 (30% ~ 85%)
                for idx, file_path in enumerate(excel_files_to_read):
                    read_percent = 30 + int(((idx + 1) / total_docs) * 55)
                    self._set_progress(read_percent, f"문서 데이터 분석 중 ({idx + 1}/{total_docs})...")

                    ext = os.path.splitext(file_path)[1].lower()
                    if ext == ".xlsx":
                        df = pd.read_excel(file_path, engine="openpyxl")
                        df.columns = df.columns.astype(str).str.strip()
                        dataframes.append(df)
                    elif ext == ".xls":
                        df = pd.read_excel(file_path, engine="xlrd")
                        df.columns = df.columns.astype(str).str.strip()
                        dataframes.append(df)
                    elif ext == ".csv":
                        try:
                            df = pd.read_csv(file_path, encoding="utf-8-sig")
                        except UnicodeDecodeError:
                            df = pd.read_csv(file_path, encoding="cp949")
                        df.columns = df.columns.astype(str).str.strip()
                        dataframes.append(df)

                # 3단계: 병합 및 최종 엑셀 파일 저장 (85% ~ 100%)
                self._set_progress(90, "데이터 컬럼 매핑 및 엑셀 파일 생성 중...")
                merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
                merged_df.to_excel(out_path, index=False, engine="openpyxl")
                self._set_progress(100, "완료")

            self.reset_ui()
            msg = (
                f"데이터 취합이 성공적으로 완료되었습니다!\n\n"
                f"• 처리된 문서 수: 총 {total_docs}개\n"
                f"• 저장 파일명: {os.path.basename(out_path)}\n"
                f"• 저장 폴더: {os.path.dirname(out_path)}"
            )
            messagebox.showinfo("취합 완료", msg)

        except Exception as e:
            self.reset_ui()
            self._set_progress(0, "오류 발생")
            messagebox.showerror("오류 발생", f"작업 중 오류가 발생했습니다:\n{str(e)}")

    def reset_ui(self):
        self.run_btn.configure(state="normal", text="데이터 취합 시작하기")

if __name__ == "__main__":
    app = ReportMergerApp()
    app.mainloop()
