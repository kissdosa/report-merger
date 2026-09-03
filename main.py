import os
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import pyzipper

# macOS 스타일 테마 설정
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

class ReportMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("EPP 보고서 자동 통합 도구")
        self.geometry("640 x 580")
        self.configure(fg_color="#F5F5F7")  # Apple Light Gray background

        self.selected_files = []
        self._init_ui()

    def _init_ui(self):
        # 상단 헤더 카드
        self.header_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=16)
        self.header_card.pack(fill="x", padx=28, pady=(24, 16))

        self.title_label = ctk.CTkLabel(
            self.header_card,
            text="보고서 데이터 병합",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=22, weight="bold"),
            text_color="#1D1D1F"
        )
        self.title_label.pack(anchor="w", padx=20, pady=(18, 4))

        self.desc_label = ctk.CTkLabel(
            self.header_card,
            text="동일한 컬럼을 가진 압축 파일(.zip) 내 보고서들을 안전하게 하나로 취합합니다.",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13),
            text_color="#86868B"
        )
        self.desc_label.pack(anchor="w", padx=20, pady=(0, 18))

        # 메인 콘텐츠 카드
        self.content_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=16)
        self.content_card.pack(fill="both", expand=True, padx=28, pady=(0, 16))

        # 1. 파일 선택 섹션
        self.file_btn = ctk.CTkButton(
            self.content_card,
            text="압축 파일(.zip) 선택",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=14, weight="bold"),
            fg_color="#0071E3",
            hover_color="#0077ED",
            corner_radius=10,
            height=38,
            command=self.select_files
        )
        self.file_btn.pack(fill="x", padx=20, pady=(20, 10))

        self.file_listbox = tk.Listbox(
            self.content_card,
            height=6,
            bd=0,
            bg="#F5F5F7",
            fg="#1D1D1F",
            selectbackground="#E5E5EA",
            selectforeground="#000000",
            font=("Apple SD Gothic Neo", 11)
        )
        self.file_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        # 2. 저장 파일명 설정
        self.name_label = ctk.CTkLabel(
            self.content_card,
            text="저장할 엑셀 파일명",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            text_color="#1D1D1F"
        )
        self.name_label.pack(anchor="w", padx=20, pady=(0, 6))

        self.filename_entry = ctk.CTkEntry(
            self.content_card,
            placeholder_text="통합_보고서.xlsx",
            corner_radius=10,
            height=38,
            fg_color="#F5F5F7",
            border_color="#D2D2D7",
            border_width=1,
            text_color="#1D1D1F"
        )
        self.filename_entry.insert(0, "통합_보고서.xlsx")
        self.filename_entry.pack(fill="x", padx=20, pady=(0, 20))

        # 실행 버튼
        self.run_btn = ctk.CTkButton(
            self,
            text="데이터 취합 및 파일 생성",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=15, weight="bold"),
            fg_color="#34C759",
            hover_color="#2DB84D",
            corner_radius=12,
            height=46,
            command=self.start_merge_thread
        )
        self.run_btn.pack(fill="x", padx=28, pady=(0, 24))

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Zip files", "*.zip")])
        if files:
            self.selected_files = list(files)
            self.file_listbox.delete(0, tk.END)
            for f in self.selected_files:
                self.file_listbox.insert(tk.END, os.path.basename(f))

    def ask_zip_password(self, filename):
        dialog = ctk.CTkInputDialog(
            text=f"[{os.path.basename(filename)}]\n비밀번호가 잠겨있습니다.\n압축 해제를 위한 번호를 입력해 주세요.",
            title="비밀번호 필요"
        )
        return dialog.get_input()

    def start_merge_thread(self):
        if not self.selected_files:
            messagebox.showwarning("안내", "병합할 압축 파일(.zip)을 하나 이상 선택해 주세요.")
            return

        out_name = self.filename_entry.get().strip()
        if not out_name:
            out_name = "통합_보고서.xlsx"
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"

        save_dir = filedialog.askdirectory(title="취합된 파일을 저장할 폴더 선택")
        if not save_dir:
            return

        out_path = os.path.join(save_dir, out_name)
        self.run_btn.configure(state="disabled", text="데이터 취합 중...")

        threading.Thread(target=self.process_merge, args=(out_path,), daemon=True).start()

    def process_merge(self, out_path):
        dataframes = []

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                for zip_path in self.selected_files:
                    extracted = False
                    password = None

                    while not extracted:
                        try:
                            with pyzipper.AESZipFile(zip_path) as zf:
                                if password:
                                    zf.setpassword(password.encode("utf-8"))
                                zf.extractall(path=temp_dir)
                            extracted = True
                        except (RuntimeError, pyzipper.BadZipFile):
                            password = self.ask_zip_password(zip_path)
                            if not password:
                                self.reset_ui()
                                messagebox.showinfo("취소", f"[{os.path.basename(zip_path)}] 암호 입력을 취소하여 작업을 중단합니다.")
                                return

                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        file_path = os.path.join(root, file)

                        if ext == ".xlsx":
                            df = pd.read_excel(file_path, engine="openpyxl")
                            dataframes.append(df)
                        elif ext == ".xls":
                            df = pd.read_excel(file_path, engine="xlrd")
                            dataframes.append(df)
                        elif ext == ".csv":
                            try:
                                df = pd.read_csv(file_path, encoding="utf-8-sig")
                            except UnicodeDecodeError:
                                df = pd.read_csv(file_path, encoding="cp949")
                            dataframes.append(df)

                if not dataframes:
                    self.reset_ui()
                    messagebox.showwarning("경고", "압축 파일 내에서 엑셀 또는 CSV 문서를 찾을 수 없습니다.")
                    return

                # 동일 컬럼 기준 병합
                merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
                merged_df.to_excel(out_path, index=False, engine="openpyxl")

            self.reset_ui()
            messagebox.showinfo("완료", f"데이터 취합이 완료되었습니다.\n저장 경로: {out_path}")

        except Exception as e:
            self.reset_ui()
            messagebox.showerror("오류 발생", f"작업 중 오류가 발생했습니다:\n{str(e)}")

    def reset_ui(self):
        self.run_btn.configure(state="normal", text="데이터 취합 및 파일 생성")

if __name__ == "__main__":
    app = ReportMergerApp()
    app.mainloop()
