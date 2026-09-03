import sys
import os
import re

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


def has_korean(text):
    """문자열 내 한글 포함 여부 검사"""
    return bool(re.search(r'[\uac00-\ud7a3\u1100-\u11ff\u3130-\u318f]', text))


def excel_col_to_index(col_str):
    """'A' -> 0, 'B' -> 1, 'Q' -> 16, 'AA' -> 26 등 엑셀 열 문자를 0-based 인덱스로 변환"""
    clean = col_str.strip().upper().replace("열", "")
    num = 0
    for c in clean:
        if "A" <= c <= "Z":
            num = num * 26 + (ord(c) - ord("A") + 1)
        else:
            return None
    return (num - 1) if num > 0 else None


# ==========================================
# 3. 디자인 가이드 적용 모달 팝업들
# ==========================================
class AboutDialog(ctk.CTkToplevel):
    """우측 상단 제품 정보 안내 팝업"""
    def __init__(self, parent, app_font):
        super().__init__(parent)

        self.title("제품 정보")
        self.geometry("380x250")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 380) // 2}+{py + (ph - 250) // 2}")

        self.header_label = ctk.CTkLabel(
            self,
            text="제품 정보",
            font=ctk.CTkFont(family=app_font, size=19, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(24, 12))

        info_frame = ctk.CTkFrame(self, fg_color="#F9FAFB", corner_radius=10)
        info_frame.pack(fill="x", padx=26, pady=(0, 18))

        info_text = (
            "• 프로그램: 엑셀 문서 자동 취합 프로그램\n"
            "• 최종 버전 : 1.5\n"
            "• 최종 수정 날짜 : 26.09.03\n"
            "• 제작자 : kkh"
        )
        self.info_label = ctk.CTkLabel(
            info_frame,
            text=info_text,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#333D4B",
            justify="left"
        )
        self.info_label.pack(anchor="w", padx=16, pady=12)

        self.btn_confirm = ctk.CTkButton(
            self,
            text="확인",
            width=140,
            height=40,
            font=ctk.CTkFont(family=app_font, size=13, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            text_color="#FFFFFF",
            corner_radius=10,
            command=self.destroy
        )
        self.btn_confirm.pack(pady=(0, 20))


class WarningDialog(ctk.CTkToplevel):
    """정렬 입력 오류 안내 모달 팝업"""
    def __init__(self, parent, message, app_font):
        super().__init__(parent)

        self.title("안내")
        self.geometry("420x220")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 420) // 2}+{py + (ph - 220) // 2}")

        self.header_label = ctk.CTkLabel(
            self,
            text="정렬 입력 안내",
            font=ctk.CTkFont(family=app_font, size=18, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(24, 10))

        self.body_label = ctk.CTkLabel(
            self,
            text=message,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#4E5968",
            justify="center",
            wraplength=360
        )
        self.body_label.pack(expand=True, padx=20, pady=(0, 16))

        self.btn_confirm = ctk.CTkButton(
            self,
            text="확인",
            width=140,
            height=42,
            font=ctk.CTkFont(family=app_font, size=13, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            text_color="#FFFFFF",
            corner_radius=10,
            command=self.destroy
        )
        self.btn_confirm.pack(pady=(0, 20))


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

        self.header_label = ctk.CTkLabel(
            self,
            text="동일 파일명 확인",
            font=ctk.CTkFont(family=app_font, size=19, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(26, 8))

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
    """작업 완료 안내 모달 팝업 (본문 내용 좌측 정렬 적용)"""
    def __init__(self, parent, total_docs, out_path, app_font, sort_info=None):
        super().__init__(parent)

        self.title("취합 완료")
        self.geometry("420x300")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 420) // 2}+{py + (ph - 300) // 2}")

        self.header_label = ctk.CTkLabel(
            self,
            text="취합 완료",
            font=ctk.CTkFont(family=app_font, size=19, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(22, 6))

        # 좌측 정렬을 위한 내부 정보 컨테이너 프레임
        detail_frame = ctk.CTkFrame(self, fg_color="#F9FAFB", corner_radius=10)
        detail_frame.pack(fill="x", padx=24, pady=(0, 16))

        lines = [
            f"• 처리된 문서 수: 총 {total_docs}개",
            f"• 정렬 기준: {sort_info if sort_info else '없음 (원본 순서)'}",
            f"• 저장 파일: {os.path.basename(out_path)}",
            f"• 저장 위치: {os.path.dirname(out_path)}"
        ]
        detail_text = "\n".join(lines)

        self.body_label = ctk.CTkLabel(
            detail_frame,
            text=detail_text,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#333D4B",
            justify="left",
            wraplength=350
        )
        self.body_label.pack(anchor="w", padx=16, pady=12)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 20))

        self.btn_folder = ctk.CTkButton(
            btn_frame,
            text="저장 폴더 열기",
            width=175,
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
            width=175,
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

        self.title("엑셀 문서 자동 취합 프로그램")
        
        self.geometry("540 x 860")
        self.minsize(480, 740)
        self.resizable(True, True)
        self.configure(fg_color="#F2F4F6")

        self.selected_files = []
        
        self.save_dir_path = str(Path.home() / "Desktop")
        if not os.path.exists(self.save_dir_path):
            self.save_dir_path = str(Path.home())

        self._init_ui()
        self._setup_drag_and_drop()

    def _init_ui(self):
        # 1. 상단 타이틀 카드 (우측 상단 제품 정보 버튼 배치)
        self.header_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=16)
        self.header_card.pack(fill="x", padx=16, pady=(16, 8))

        top_bar = ctk.CTkFrame(self.header_card, fg_color="transparent")
        top_bar.pack(fill="x", padx=20, pady=(16, 4))

        self.title_label = ctk.CTkLabel(
            top_bar,
            text="엑셀 문서 자동 취합 프로그램",
            font=ctk.CTkFont(family=APP_FONT, size=20, weight="bold"),
            text_color="#191F28"
        )
        self.title_label.pack(side="left")

        self.btn_about = ctk.CTkButton(
            top_bar,
            text="제품 정보",
            width=75,
            height=30,
            font=ctk.CTkFont(family=APP_FONT, size=11),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#4E5968",
            corner_radius=8,
            command=self.open_about_dialog
        )
        self.btn_about.pack(side="right")

        guide_text = "• 다수의 엑셀 문서 및 압축 파일(.zip) 내 데이터를 컬럼 기준으로 자동 병합합니다."
        self.desc_label = ctk.CTkLabel(
            self.header_card,
            text=guide_text,
            font=ctk.CTkFont(family=APP_FONT, size=12),
            text_color="#4E5968",
            justify="left"
        )
        self.desc_label.pack(anchor="w", padx=20, pady=(0, 16))

        # 2. 메인 콘텐츠 카드
        self.content_card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=16)
        self.content_card.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # 파일 선택 및 목록 비우기
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

        # 파일 목록 박스 (드래그 앤 드롭 지원)
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
        self.file_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 10))

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
            height=36,
            fg_color="#F9FAFB",
            border_color="#E5E8EB",
            border_width=1,
            text_color="#191F28"
        )
        self.filename_entry.insert(0, "통합_보고서.xlsx")
        self.filename_entry.pack(fill="x", padx=20, pady=(0, 10))

        # 저장 경로
        self.path_label = ctk.CTkLabel(
            self.content_card,
            text="저장 위치 (기본값: 바탕화면)",
            font=ctk.CTkFont(family=APP_FONT, size=12, weight="bold"),
            text_color="#333D4B"
        )
        self.path_label.pack(anchor="w", padx=20, pady=(0, 4))

        self.path_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.path_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.path_entry = ctk.CTkEntry(
            self.path_frame,
            font=ctk.CTkFont(family=APP_FONT, size=12),
            corner_radius=10,
            height=36,
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
            height=36,
            font=ctk.CTkFont(family=APP_FONT, size=12),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=10,
            command=self.change_directory
        )
        self.path_btn.pack(side="right")

        # 데이터 정렬 옵션 영역
        self.sort_label = ctk.CTkLabel(
            self.content_card,
            text="데이터 정렬 옵션 (선택 사항)",
            font=ctk.CTkFont(family=APP_FONT, size=12, weight="bold"),
            text_color="#333D4B"
        )
        self.sort_label.pack(anchor="w", padx=20, pady=(0, 4))

        self.sort_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.sort_frame.pack(fill="x", padx=20, pady=(0, 2))

        self.sort_col_var = tk.StringVar()
        self.sort_col_var.trace_add("write", self._on_sort_text_change)

        self.sort_col_entry = ctk.CTkEntry(
            self.sort_frame,
            textvariable=self.sort_col_var,
            font=ctk.CTkFont(family=APP_FONT, size=12),
            placeholder_text="기준 열 (예: Q, B, AA 등 영문)",
            corner_radius=10,
            height=36,
            fg_color="#F9FAFB",
            border_color="#E5E8EB",
            border_width=1,
            text_color="#191F28"
        )
        self.sort_col_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.sort_order_menu = ctk.CTkOptionMenu(
            self.sort_frame,
            values=["정렬 안 함", "오름차순 (1→9, A→Z)", "내림차순 (9→1, Z→A)"],
            font=ctk.CTkFont(family=APP_FONT, size=11),
            dropdown_font=ctk.CTkFont(family=APP_FONT, size=11),
            fg_color="#F2F4F6",
            button_color="#E5E8EB",
            button_hover_color="#D2D6DB",
            text_color="#333D4B",
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color="#191F28",
            corner_radius=10,
            width=165,
            height=36
        )
        self.sort_order_menu.set("정렬 안 함")
        self.sort_order_menu.pack(side="right")

        # 실시간 한글 입력 시 표시되는 붉은 경고 문구 라벨
        self.sort_warn_label = ctk.CTkLabel(
            self.content_card,
            text="",
            font=ctk.CTkFont(family=APP_FONT, size=11, weight="bold"),
            text_color="#FF3B30"
        )
        self.sort_warn_label.pack(anchor="w", padx=22, pady=(0, 8))

        # 프로그레스 영역
        self.progress_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=20, pady=(0, 10))

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

        # 4. 맨 하단 면책 문구
        disclaimer_text = (
            "본 프로그램은 업무 지원용 도구로 자유로운 수정, 사용 및 배포가 가능합니다.\n"
            "(사용 중 발생하는 모든 문제는 사용자 본인의 책임입니다.)"
        )
        self.disclaimer_label = ctk.CTkLabel(
            self,
            text=disclaimer_text,
            font=ctk.CTkFont(family=APP_FONT, size=11),
            text_color="#8B95A1",
            justify="center"
        )
        self.disclaimer_label.pack(padx=18, pady=(0, 14))

    def _on_sort_text_change(self, *args):
        """정렬 입력창 텍스트 변경 감지 -> 한글 입력 시 실시간 붉은 경고 표시"""
        val = self.sort_col_var.get()
        if has_korean(val):
            self.sort_warn_label.configure(text="영문으로 입력해 주세요")
        else:
            self.sort_warn_label.configure(text="")

    def open_about_dialog(self):
        AboutDialog(self, APP_FONT)

    def _setup_drag_and_drop(self):
        """Tkinter 메인 루프 충돌 방지 안전한 드래그 앤 드롭 연동"""
        try:
            import windnd

            def on_drop(files):
                # UI 스레드로 안전하게 비동기 전달 (크래시 원천 방지)
                self.after(10, self._process_dropped_files, files)

            windnd.hook_dropfiles(self.file_listbox, func=on_drop)
        except Exception:
            pass

    def _process_dropped_files(self, file_paths):
        try:
            added_count = 0
            for path in file_paths:
                if isinstance(path, bytes):
                    try:
                        path = path.decode("utf-8")
                    except UnicodeDecodeError:
                        path = path.decode("cp949", errors="ignore")

                ext = os.path.splitext(path)[1].lower()
                if ext in [".zip", ".xlsx", ".xls", ".csv"]:
                    if path not in self.selected_files:
                        self.selected_files.append(path)
                        added_count += 1

            if added_count > 0:
                self._refresh_listbox()
        except Exception:
            pass

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

        # 한글 입력 여부 엄격 검사 -> 경고 팝업 발생 후 중단
        sort_col_input = self.sort_col_entry.get().strip()
        if sort_col_input and has_korean(sort_col_input):
            warn_msg = "데이터 정렬 값을 영문으로 표시해 주세요.\n(A열을 정렬하고 싶으면 A, B열을 정렬하고 싶으면 B 등)"
            WarningDialog(self, warn_msg, APP_FONT)
            self.sort_col_entry.focus()
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

        if os.path.exists(out_path):
            dlg = OverwriteDialog(self, out_name, APP_FONT)
            self.wait_window(dlg)
            if not dlg.result:
                self.filename_entry.focus()
                return

        sort_order_mode = self.sort_order_menu.get()

        self.run_btn.configure(
            state="disabled",
            text="데이터 취합 중입니다. 잠시만 기다려주세요.",
            fg_color="#F2F4F6",
            text_color="#8B95A1"
        )
        self._set_progress(0, "취합 준비 중...")

        threading.Thread(
            target=self.process_merge,
            args=(out_path, sort_col_input, sort_order_mode),
            daemon=True
        ).start()

    def process_merge(self, out_path, sort_col_input, sort_order_mode):
        dataframes = []
        excel_files_to_read = []
        sort_summary = None

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

                # 2단계: 데이터 로드 (30% ~ 80%)
                for idx, file_path in enumerate(excel_files_to_read):
                    read_percent = 30 + int(((idx + 1) / total_docs) * 50)
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

                # 3단계: 통합 및 데이터 정렬 (80% ~ 100%)
                self._set_progress(85, "데이터 컬럼 매핑 중...")
                merged_df = pd.concat(dataframes, ignore_index=True, sort=False)

                # 정렬 로직 적용
                if sort_col_input and sort_order_mode != "정렬 안 함":
                    self._set_progress(90, "설정된 기준 열로 데이터 정렬 중...")
                    target_col = None

                    col_idx = excel_col_to_index(sort_col_input)
                    if col_idx is not None and 0 <= col_idx < len(merged_df.columns):
                        target_col = merged_df.columns[col_idx]
                    elif sort_col_input in merged_df.columns:
                        target_col = sort_col_input

                    if target_col is not None:
                        is_ascending = "오름차순" in sort_order_mode
                        try:
                            merged_df = merged_df.sort_values(
                                by=target_col,
                                ascending=is_ascending,
                                na_position="last"
                            )
                        except Exception:
                            merged_df = merged_df.sort_values(
                                by=target_col,
                                ascending=is_ascending,
                                na_position="last",
                                key=lambda s: s.astype(str)
                            )
                        sort_summary = f"{sort_col_input.upper()}열({target_col}) 기준 {'오름차순' if is_ascending else '내림차순'}"

                self._set_progress(95, "최종 엑셀 파일 생성 중...")
                merged_df.to_excel(out_path, index=False, engine="openpyxl")
                self._set_progress(100, "완료")

            self.reset_ui()
            # 커스텀 취합 완료 팝업 호출 (좌측 정렬 상세 정보)
            self.after(0, lambda: CompleteDialog(self, total_docs, out_path, APP_FONT, sort_summary))

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
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass

    app = ReportMergerApp()
    app.mainloop()
