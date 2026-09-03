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


class OverwriteDialog(ctk.CTkToplevel):
    """동일 파일명 존재 시 덮어쓰기 여부를 묻는 팝업 다이얼로그"""
    def __init__(self, parent, filename):
        super().__init__(parent)
        self.result = False

        self.title("중복 파일 확인")
        self.geometry("420x210")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        # 모달 창 설정
        self.transient(parent)
        self.grab_set()

        # 화면 중앙 배치
        self.update_idletasks()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        x = parent_x + (parent_w - 420) // 2
        y = parent_y + (parent_h - 210) // 2
        self.geometry(f"+{x}+{y}")

        # 안내 문구
        self.msg_label = ctk.CTkLabel(
            self,
            text=f"저장 경로에 [{filename}] 파일이 이미 존재합니다.\n\n동일한 파일명이 있습니다. 덮어 씌우시겠습니까?",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            text_color="#191F28",
            justify="center",
            wraplength=380
        )
        self.msg_label.pack(expand=True, padx=20, pady=(24, 16))

        # 버튼 영역
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 24))

        # 1. 예 (덮어쓰기)
        self.btn_yes = ctk.CTkButton(
            btn_frame,
            text="예",
            width=170,
            height=42,
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            corner_radius=10,
            command=self._on_yes
        )
        self.btn_yes.pack(side="left", padx=(0, 10))

        # 2. 파일명을 새로 지정 (취소)
        self.btn_cancel = ctk.CTkButton(
            btn_frame,
            text="파일명을 새로 지정",
            width=170,
            height=42,
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=10,
            command=self._on_cancel
        )
        self.btn_cancel.pack(side="right")

    def _on_yes(self):
        self.result = True
        self.destroy()

    def _on_cancel(self):
        self.result = False
        self.destroy()


class ReportMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("문서 통합 프로그램")
        self.geometry("630 x 800")
        self.resizable(False, False)
        
        # Apple / Toss 시그니처 연회색 배경
        self.configure(fg_color="#F2F4F6")

        self.selected_files = []
        
        # 기본 저장 경로: 사용자 바탕화면
        self.save_dir_path = str(Path.home() / "Desktop")
        if not os.path.exists(self.save_dir_path):
            self.save_dir_path = str(Path.home())

        self._init_ui()

    def _init_ui(self):
        # 1. 상단 타이틀 및 안내 카드
        self.header_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=18)
        self.header_card.pack(fill="x", padx=20, pady=(18, 10))

        self.title_label = ctk.CTkLabel(
            self.header_card,
            text="문서 통합 프로그램",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=21, weight="bold"),
            text_color="#191F28"
        )
        self.title_label.pack(anchor="w", padx=22, pady=(18, 8))

        guide_text = (
            "• 다수의 엑셀 문서 및 압축 파일(.zip) 내 데이터를 컬럼 기준으로 자동 병합합니다.\n"
            "• 문서 간 컬럼이 서로 다른 경우, 전체 컬럼 항목을 기준으로 일치하는 데이터를 누락 없이 정렬하여 취합합니다."
        )
        self.desc_label = ctk.CTkLabel(
            self.header_card,
            text=guide_text,
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12),
            text_color="#4E5968",
            justify="left"
        )
        self.desc_label.pack(anchor="w", padx=22, pady=(0, 18))

        # 2. 메인 설정 카드
        self.content_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=18)
        self.content_card.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        # [섹션 1] 파일 선택 버튼 영역
        self.btn_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=22, pady=(18, 8))

        self.file_btn = ctk.CTkButton(
            self.btn_frame,
            text="+ 취합할 파일 추가 (.zip, .xlsx, .csv)",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            fg_color="#3182F6",
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
        self.file_listbox.pack(fill="both", expand=True, padx=22, pady=(0, 14))

        # [섹션 2] 저장할 파일명
        self.name_label = ctk.CTkLabel(
            self.content_card,
            text="저장할 엑셀 파일 이름",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            text_color="#333D4B"
        )
        self.name_label.pack(anchor="w", padx=22, pady=(0, 5))

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
        self.filename_entry.pack(fill="x", padx=22, pady=(0, 14))

        # [섹션 3] 저장 경로
        self.path_label = ctk.CTkLabel(
            self.content_card,
            text="저장 위치 (기본값: 바탕화면)",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            text_color="#333D4B"
        )
        self.path_label.pack(anchor="w", padx=22, pady=(0, 5))

        self.path_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.path_frame.pack(fill="x", padx=22, pady=(0, 14))

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
        self.progress_frame.pack(fill="x", padx=22, pady=(0, 14))

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
        self.progress_bar.pack(fill="x", padx=22, pady=(0, 16))

        # 3. 하단 액션 버튼
        self.run_btn = ctk.CTkButton(
            self,
            text="문서 취합 시작하기",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=15, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            corner_radius=14,
            height=48,
            command=self.start_merge_thread
        )
        self.run_btn.pack(fill="x", padx=20, pady=(0, 10))

        # 4. 맨 하단 정보 영역 (면책 문구 및 버전/제작자 정보)
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.footer_frame.pack(fill="x", padx=22, pady=(0, 14))

        # 좌측: 면책 문구
        disclaimer_text = (
            "본 프로그램은 업무 지원용 도구로 자유로운 수정, 사용 및 배포가 가능합니다.\n"
            "(사용 중 발생하는 모든 문제는 사용자 본인의 책임입니다.)"
        )
        self.disclaimer_label = ctk.CTkLabel(
            self.footer_frame,
            text=disclaimer_text,
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=10),
            text_color="#8B95A1",
            justify="left"
        )
        self.disclaimer_label.pack(side="left")

        # 우측: 최종 정보 (v1.3, 날짜, 제작자)
        meta_text = "최종 버전 : 1.3\n최종 수정 날짜 : 26.09.03\n제작자 : kkh"
        self.meta_label = ctk.CTkLabel(
            self.footer_frame,
            text=meta_text,
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=10),
            text_color="#8B95A1",
            justify="right"
        )
        self.meta_label.pack(side="right")

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

        # 동일한 파일명이 이미 존재하는지 확인
        if os.path.exists(out_path):
            dlg = OverwriteDialog(self, out_name)
            self.wait_window(dlg)
            if not dlg.result:
                # '파일명을 새로 지정' 선택 시 입력창에 포커스 후 중단
                self.filename_entry.focus()
                return

        # 버튼 문구 변경 (따옴표 없음)
        self.run_btn.configure(state="disabled", text="데이터 취합 중입니다. 잠시만 기다려주세요.")
        self._set_progress(0, "취합 준비 중...")

        threading.Thread(target=self.process_merge, args=(out_path,), daemon=True).start()

    def process_merge(self, out_path):
        dataframes = []
        excel_files_to_read = []

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_files = [f for f in self.selected_files if os.path.splitext(f)[1].lower() == ".zip"]
                direct_files = [f for f in self.selected_files if os.path.splitext(f)[1].lower() in [".xlsx", ".xls", ".csv"]]

                # 1단계: 압축 해제 (0% ~ 30%)
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

                # 2단계: 데이터프레임 읽기 (30% ~ 85%)
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

                # 3단계: 동일 컬럼 기준 병합 및 저장 (85% ~ 100%)
                self._set_progress(90, "데이터 컬럼 매핑 및 엑셀 파일 생성 중...")
                merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
                merged_df.to_excel(out_path, index=False, engine="openpyxl")
                self._set_progress(100, "완료")

            self.reset_ui()
            msg = (
                f"문서 취합이 성공적으로 완료되었습니다!\n\n"
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
        self.run_btn.configure(state="normal", text="문서 취합 시작하기")


if __name__ == "__main__":
    app = ReportMergerApp()
    app.mainloop()
