import os
import tempfile
import threading
from pathlib import Path
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
        self.geometry("640 x 680")
        self.resizable(False, False)
        self.configure(fg_color="#F5F5F7")  # Apple Light Gray background

        self.selected_files = []
        
        # 디폴트 저장 경로: 사용자 바탕화면
        self.save_dir_path = str(Path.home() / "Desktop")
        if not os.path.exists(self.save_dir_path):
            self.save_dir_path = str(Path.home())

        self._init_ui()

    def _init_ui(self):
        # 1. 상단 헤더 카드
        self.header_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=14)
        self.header_card.pack(fill="x", padx=24, pady=(20, 12))

        self.title_label = ctk.CTkLabel(
            self.header_card,
            text="보고서 데이터 병합",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=20, weight="bold"),
            text_color="#1D1D1F"
        )
        self.title_label.pack(anchor="w", padx=20, pady=(16, 4))

        self.desc_label = ctk.CTkLabel(
            self.header_card,
            text="다수의 압축 파일(.zip) 내 엑셀 데이터를 컬럼 기준으로 자동 병합합니다.",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13),
            text_color="#86868B"
        )
        self.desc_label.pack(anchor="w", padx=20, pady=(0, 16))

        # 2. 메인 콘텐츠 카드
        self.content_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=14)
        self.content_card.pack(fill="both", expand=True, padx=24, pady=(0, 12))

        # [섹션 1] 압축 파일 선택
        self.file_btn = ctk.CTkButton(
            self.content_card,
            text="압축 파일(.zip) 다중 선택",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=13, weight="bold"),
            fg_color="#0071E3",
            hover_color="#0077ED",
            corner_radius=8,
            height=36,
            command=self.select_files
        )
        self.file_btn.pack(fill="x", padx=20, pady=(16, 8))

        self.file_listbox = tk.Listbox(
            self.content_card,
            height=5,
            bd=0,
            bg="#F5F5F7",
            fg="#1D1D1F",
            highlightthickness=1,
            highlightbackground="#E5E5EA",
            selectbackground="#D2D2D7",
            selectforeground="#000000",
            font=("Apple SD Gothic Neo", 10)
        )
        self.file_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        # [섹션 2] 저장할 파일명 설정
        self.name_label = ctk.CTkLabel(
            self.content_card,
            text="저장할 엑셀 파일명",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12, weight="bold"),
            text_color="#1D1D1F"
        )
        self.name_label.pack(anchor="w", padx=20, pady=(0, 4))

        self.filename_entry = ctk.CTkEntry(
            self.content_card,
            placeholder_text="통합_보고서.xlsx",
            corner_radius=8,
            height=36,
            fg_color="#F5F5F7",
            border_color="#D2D2D7",
            border_width=1,
            text_color="#1D1D1F"
        )
        self.filename_entry.insert(0, "통합_보고서.xlsx")
        self.filename_entry.pack(fill="x", padx=20, pady=(0, 12))

        # [섹션 3] 저장 경로 설정 (디폴트: 바탕화면)
        self.path_label = ctk.CTkLabel(
            self.content_card,
            text="저장 경로 (기본값: 바탕화면)",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12, weight="bold"),
            text_color="#1D1D1F"
        )
        self.path_label.pack(anchor="w", padx=20, pady=(0, 4))

        self.path_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.path_frame.pack(fill="x", padx=20, pady=(0, 16))

        self.path_entry = ctk.CTkEntry(
            self.path_frame,
            corner_radius=8,
            height=36,
            fg_color="#F5F5F7",
            border_color="#D2D2D7",
            border_width=1,
            text_color="#1D1D1F"
        )
        self.path_entry.insert(0, self.save_dir_path)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.path_btn = ctk.CTkButton(
            self.path_frame,
            text="경로 변경",
            width=80,
            height=36,
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=12),
            fg_color="#E5E5EA",
            hover_color="#D2D2D7",
            text_color="#1D1D1F",
            corner_radius=8,
            command=self.change_directory
        )
        self.path_btn.pack(side="right")

        # 3. 하단 실행 버튼
        self.run_btn = ctk.CTkButton(
            self,
            text="데이터 취합 및 파일 생성",
            font=ctk.CTkFont(family="Apple SD Gothic Neo", size=15, weight="bold"),
            fg_color="#34C759",
            hover_color="#2DB84D",
            corner_radius=10,
            height=46,
            command=self.start_merge_thread
        )
        self.run_btn.pack(fill="x", padx=24, pady=(0, 20))

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="취합할 압축 파일(.zip)을 선택하세요 (복수 선택 가능)",
            filetypes=[("Zip files", "*.zip")]
        )
        if files:
            self.selected_files = list(files)
            self.file_listbox.delete(0, tk.END)
            for f in self.selected_files:
                self.file_listbox.insert(tk.END, f"  {os.path.basename(f)}")

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

    def start_merge_thread(self):
        if not self.selected_files:
            messagebox.showwarning("안내", "병합할 압축 파일(.zip)을 먼저 선택해 주세요.")
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
        self.run_btn.configure(state="disabled", text="데이터 취합 중...")

        threading.Thread(target=self.process_merge, args=(out_path,), daemon=True).start()

    def process_merge(self, out_path):
        dataframes = []
        file_count = 0

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # 덮어쓰기 방지를 위해 각 zip 파일마다 독립된 하위 폴더에 압축 해제
                for idx, zip_path in enumerate(self.selected_files):
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

                # 모든 하위 폴더에서 엑셀 및 CSV 파일 수집
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        file_path = os.path.join(root, file)

                        if ext == ".xlsx":
                            df = pd.read_excel(file_path, engine="openpyxl")
                            df.columns = df.columns.astype(str).str.strip()
                            dataframes.append(df)
                            file_count += 1
                        elif ext == ".xls":
                            df = pd.read_excel(file_path, engine="xlrd")
                            df.columns = df.columns.astype(str).str.strip()
                            dataframes.append(df)
                            file_count += 1
                        elif ext == ".csv":
                            try:
                                df = pd.read_csv(file_path, encoding="utf-8-sig")
                            except UnicodeDecodeError:
                                df = pd.read_csv(file_path, encoding="cp949")
                            df.columns = df.columns.astype(str).str.strip()
                            dataframes.append(df)
                            file_count += 1

                if not dataframes:
                    self.reset_ui()
                    messagebox.showwarning("경고", "선택한 압축 파일 내부에서 엑셀 또는 CSV 문서를 찾을 수 없습니다.")
                    return

                # 동일 컬럼 기준 자동 매핑 병합 (불일치 컬럼은 확장 매핑)
                merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
                merged_df.to_excel(out_path, index=False, engine="openpyxl")

            self.reset_ui()
            messagebox.showinfo(
                "저장 완료",
                f"총 {len(self.selected_files)}개의 압축 파일(보고서 {file_count}개) 취합이 완료되었습니다!\n\n저장 경로:\n{out_path}"
            )

        except Exception as e:
            self.reset_ui()
            messagebox.showerror("오류 발생", f"작업 중 오류가 발생했습니다:\n{str(e)}")

    def reset_ui(self):
        self.run_btn.configure(state="normal", text="데이터 취합 및 파일 생성")

if __name__ == "__main__":
    app = ReportMergerApp()
    app.mainloop()
