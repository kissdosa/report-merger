import sys
import os

# ==========================================
# 1. 중복 실행 방지 (Windows Mutex)
# ==========================================
if sys.platform == "win32":
    import ctypes
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "ReportMerger_Mutex_SingleInstance")
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        sys.exit(0)

# ==========================================
# 2. 메인 모듈 로드
# ==========================================
import tempfile
import threading
from pathlib import Path
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import pyzipper

# 외관 테마 설정
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


def get_system_font():
    """나눔고딕 -> 맑은 고딕 -> 시스템 기본 고딕 순서로 감지"""
    dummy = tk.Tk()
    dummy.withdraw()
    fonts = [f.lower() for f in tkfont.families(dummy)]
    dummy.destroy()

    if any("nanumgothic" in f or "나눔고딕" in f for f in fonts):
        return "NanumGothic"
    elif any("malgun" in f or "맑은 고딕" in f for f in fonts):
        return "Malgun Gothic"
    else:
        return "Segoe UI"


APP_FONT = get_system_font()


# ==========================================
# 3. 디자인 가이드 적용 팝업 모달
# ==========================================
class OverwriteDialog(ctk.CTkToplevel):
    """동일 파일명 존재 시 덮어쓰기 여부를 묻는 모달 팝업"""
    def __init__(self, parent, filename, app_font):
        super().__init__(parent)
        self.result = False

        self.title("중복 파일 확인")
        self.geometry("400x240")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 400) // 2}+{py + (ph - 240) // 2}")

        # 1. Header Text
        self.header_label = ctk.CTkLabel(
            self,
            text="동일 파일명 확인",
            font=ctk.CTkFont(family=app_font, size=19, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(26, 8))

        # 2. Body Text
        body_text = (
            f"저장 위치에 [{filename}] 파일이 이미 존재합니다.\n\n"
            "기존 파일에 덮어 씌우시겠습니까?"
        )
        self.body_label = ctk.CTkLabel(
            self,
            text=body_text,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#4E5968",
            justify="center",
            wraplength=340
        )
        self.body_label.pack(expand=True, padx=20, pady=(0, 18))

        # 3. Bottom Action Buttons (좌: 보조 / 우: 주요 실행)
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 22))

        self.btn_cancel = ctk.CTkButton(
            btn_frame,
            text="파일명을 새로 지정",
            width=165,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=12,
            command=self._on_cancel
        )
        self.btn_cancel.pack(side="left", padx=(0, 8))

        self.btn_yes = ctk.CTkButton(
            btn_frame,
            text="예",
            width=165,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            text_color="#FFFFFF",
            corner_radius=12,
            command=self._on_yes
        )
        self.btn_yes.pack(side="right")

    def _on_yes(self):
        self.result = True
        self.destroy()

    def _on_cancel(self):
        self.result = False
        self.destroy()


class CompleteDialog(ctk.CTkToplevel):
    """작업 완료 안내 모달 팝업"""
    def __init__(self, parent, total_docs, out_path, app_font):
        super().__init__(parent)

        self.title("취합 완료")
        self.geometry("400x270")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 400) // 2}+{py + (ph - 270) // 2}")

        # 1. Header Text
        self.header_label = ctk.CTkLabel(
            self,
            text="취합 완료",
            font=ctk.CTkFont(family=app_font, size=19, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(26, 8))

        # 2. Body Text
        body_text = (
            "문서 취합이 성공적으로 완료되었습니다.\n\n"
            f"• 처리된 문서 수: 총 {total_docs}개\n"
            f"• 저장 파일: {os.path.basename(out_path)}\n"
            f"• 저장 위치: {os.path.dirname(out_path)}"
        )
        self.body_label = ctk.CTkLabel(
            self,
            text=body_text,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#4E5968",
            justify="center",
            wraplength=340
        )
        self.body_label.pack(expand=True, padx=20, pady=(0, 18))

        # 3. Bottom Action Buttons (좌: 폴더 열기 / 우: 확인)
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 22))

        self.btn_folder = ctk.CTkButton(
            btn_frame,
            text="저장 폴더 열기",
            width=165,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=12,
            command=lambda: self._open_folder(out_path)
        )
        self.btn_folder.pack(side="left", padx=(0, 8))

        self.btn_confirm = ctk.CTkButton(
            btn_frame,
            text="확인",
            width=165,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            text_color="#FFFFFF",
            corner_radius=12,
            command=self.destroy
        )
        self.btn_confirm.pack(side="right")

    def _open_folder(self, path):
        folder = os.path.dirname(path)
        if sys.platform == "win32":
            os.startfile(folder)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.run(["open", folder])
        self.destroy()


# ==========================================
# 4. 메인 애플리케이션
# ==========================================
class ReportMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("문서 통합 프로그램")
        
        # 슬림한 가로 너비(520px) 및 자유로운 크기 조절
        self.geometry("520 x 780")
        self.minsize(450, 680)
        self.resizable(True, True)
        self.configure(fg_color="#F2F4F6")

        self.selected_files = []
        
        self.save_dir_path = str(Path.home() / "Desktop")
        if not os.path.exists(self.save_dir_path):
            self.save_dir_path = str(Path.home())

        self._init_ui()

    def _init_ui(self):
        # 1. 상단 안내 카드
        self.header_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=16)
        self.header_card.pack(fill="x", padx=16, pady=(16, 8))

        self.title_label = ctk.CTkLabel(
            self.header_card,
            text="문서 통합 프로그램",
            font=ctk.CTkFont(family=APP_FONT, size=20, weight="bold"),
            text_color="#191F28"
        )
        self.title_label.pack(anchor="w", padx=20, pady=(16, 6))

        guide_text = (
            "• 다수의 엑셀 문서 및 압축 파일(.zip) 내 데이터를 컬럼 기준으로 자동 병합합니다.\n"
            "• 문서 간 컬럼이 서로 다른 경우, 전체 컬럼 항목을 기준으로 일치하는 데이터를 누락 없이 정렬하여 취합합니다."
        )
        self.desc_label = ctk.CTkLabel(
            self.header_card,
            text=guide_text,
            font=ctk.CTkFont(family=APP_FONT, size=11),
            text_color="#4E5968",
            justify="left"
        )
        self.desc_label.pack(anchor="w", padx=20, pady=(0, 16))

        # 2. 메인 콘텐츠 카드
        self.content_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=16)
        self.content_card.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # 파일 선택 버튼 영역
        self.btn_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=20, pady=(16, 6))

        self.file_btn = ctk.CTkButton(
            self.btn_frame,
            text="+ 취합할 파일 추가 (.zip, .xlsx, .csv)",
            font=ctk.CTkFont(family=APP_FONT, size=12, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            corner_radius=10,
            height=36,
            command=self.select_files
        )
        self.file_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.clear_btn = ctk.CTkButton(
            self.btn_frame,
            text="목록 비우기",
            font=ctk.CTkFont(family=APP_FONT, size=12),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#4E5968",
            corner_radius=10,
            width=80,
            height=36,
            command=self.clear_file_list
        )
        self.clear_btn.pack(side="right")

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
            font=(APP_FONT, 10)
        )
        self.file_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        # 저장 파일명
        self.name_label = ctk.CTkLabel(
            self.content_card,
            text="저장할 엑셀 파일 이름",
            font=ctk.CTkFont(family=APP_FONT, size=12, weight="bold"),
            text_color="#333D4B"
        )
        self.name_label.pack(anchor="w", padx=20, pady=(0, 4))

        self.filename_entry = ctk.CTkEntry(
            self.content_card,
            font=ctk.CTkFont(family=APP_FONT, size=12),
            placeholder_text="통합_보고서.xlsx",
            corner_radius=10,
            height=38,
            fg_color="#F9FAFB",
            border_color="#E5E8EB",
            border_width=1,
            text_color="#191F28"
        )
        self.filename_entry.insert(0, "통합_보고서.xlsx")
        self.filename_entry.pack(fill="x", padx=20, pady=(0, 12))

        # 저장 경로
        self.path_label = ctk.CTkLabel(
            self.content_card,
            text="저장 위치 (기본값: 바탕화면)",
            font=ctk.CTkFont(family=APP_FONT, size=12, weight="bold"),
            text_color="#333D4B"
        )
        self.path_label.pack(anchor="w", padx=20, pady=(0, 4))

        self.path_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.path_frame.pack(fill="x", padx=20, pady=(0, 12))

        self.path_entry = ctk.CTkEntry(
            self.path_frame,
            font=ctk.CTkFont(family=APP_FONT, size=12),
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
            width=76,
            height=38,
            font=ctk.CTkFont(family=APP_FONT, size=12),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=10,
            command=self.change_directory
        )
        self.path_btn.pack(side="right")

        # 프로그레스 영역
        self.progress_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=20, pady=(0, 12))

        self.progress_status_label = ctk.CTkLabel(
            self.progress_frame,
            text="대기 중",
            font=ctk.CTkFont(family=APP_FONT, size=11),
            text_color="#8B95A1"
        )
        self.progress_status_label.pack(side="left")

        self.progress_percent_label = ctk.CTkLabel(
            self.progress_frame,
            text="0%",
            font=ctk.CTkFont(family=APP_FONT, size=11, weight="bold"),
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
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 16))

        # 3. 하단 CTA 실행 버튼
        self.run_btn = ctk.CTkButton(
            self,
            text="문서 취합 시작하기",
            font=ctk.CTkFont(family=APP_FONT, size=15, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            text_color="#FFFFFF",
            corner_radius=12,
            height=48,
            command=self.start_merge_thread
        )
        self.run_btn.pack(fill="x", padx=16, pady=(0, 10))

        # 4. 맨 하단 정보 (면책 문구 + 메타데이터)
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.footer_frame.pack(fill="x", padx=18, pady=(0, 12))

        disclaimer_text = (
            "본 프로그램은 업무 지원용 도구로 자유로운 수정, 사용 및 배포가 가능합니다.\n"
            "(사용 중 발생하는 모든 문제는 사용자 본인의 책임입니다.)"
        )
        self.disclaimer_label = ctk.CTkLabel(
            self.footer_frame,
            text=disclaimer_text,
            font=ctk.CTkFont(family=APP_FONT, size=9),
            text_color="#8B95A1",
            justify="left"
        )
        self.disclaimer_label.pack(side="left")

        meta_text = "최종 버전 : 1.3\n최종 수정 날짜 : 26.09.03\n제작자 : kkh"
        self.meta_label = ctk.CTkLabel(
            self.footer_frame,
            text=meta_text,
            font=ctk.CTkFont(family=APP_FONT, size=9),
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

        # 동일 파일명 중복 확인 모달 호출
        if os.path.exists(out_path):
            dlg = OverwriteDialog(self, out_name, APP_FONT)
            self.wait_window(dlg)
            if not dlg.result:
                self.filename_entry.focus()
                return

        self.run_btn.configure(
            state="disabled",
            text="데이터 취합 중입니다. 잠시만 기다려주세요.",
            fg_color="#F2F4F6",
            text_color="#8B95A1"
        )
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

                # 3단계: 통합 및 엑셀 저장 (85% ~ 100%)
                self._set_progress(90, "데이터 컬럼 매핑 및 엑셀 파일 생성 중...")
                merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
                merged_df.to_excel(out_path, index=False, engine="openpyxl")
                self._set_progress(100, "완료")

            self.reset_ui()
            # 커스텀 작업 완료 모달 팝업 호출
            self.after(0, lambda: CompleteDialog(self, total_docs, out_path, APP_FONT))

        except Exception as e:
            self.reset_ui()
            self._set_progress(0, "오류 발생")
            messagebox.showerror("오류 발생", f"작업 중 오류가 발생했습니다:\n{str(e)}")

    def reset_ui(self):
        self.run_btn.configure(
            state="normal",
            text="문서 취합 시작하기",
            fg_color="#3182F6",
            text_color="#FFFFFF"
        )


if __name__ == "__main__":
    # PyInstaller 부트로더 스플래시 종료 (즉시 메인 윈도우로 전환)
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass

    app = ReportMergerApp()
    app.mainloop()
