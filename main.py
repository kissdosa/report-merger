import sys
import os
import re
import io

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
from tkinter import filedialog
import customtkinter as ctk
import pandas as pd
import pyzipper
from PIL import Image, ImageDraw

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


def extract_columns_from_file(file_path):
    """파일에서 첫 번째 행(컬럼 헤더 목록)을 안전하게 추출"""
    ext = os.path.splitext(file_path)[1].lower()
    cols = []
    try:
        if ext == ".xlsx":
            df = pd.read_excel(file_path, nrows=1, engine="openpyxl")
            cols = [str(c).strip() for c in df.columns]
        elif ext == ".xls":
            df = pd.read_excel(file_path, nrows=1, engine="xlrd")
            cols = [str(c).strip() for c in df.columns]
        elif ext == ".csv":
            try:
                df = pd.read_csv(file_path, nrows=1, encoding="utf-8-sig")
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, nrows=1, encoding="cp949")
            cols = [str(c).strip() for c in df.columns]
        elif ext == ".zip":
            with pyzipper.AESZipFile(file_path) as zf:
                for name in zf.namelist():
                    sub_ext = os.path.splitext(name)[1].lower()
                    if sub_ext in [".xlsx", ".xls", ".csv"] and not name.startswith("__MACOSX"):
                        try:
                            data = zf.read(name)
                            bio = io.BytesIO(data)
                            if sub_ext == ".xlsx":
                                df = pd.read_excel(bio, nrows=1, engine="openpyxl")
                            elif sub_ext == ".xls":
                                df = pd.read_excel(bio, nrows=1, engine="xlrd")
                            elif sub_ext == ".csv":
                                try:
                                    df = pd.read_csv(bio, nrows=1, encoding="utf-8-sig")
                                except UnicodeDecodeError:
                                    bio.seek(0)
                                    df = pd.read_csv(bio, nrows=1, encoding="cp949")
                            cols = [str(c).strip() for c in df.columns]
                            if cols:
                                break
                        except Exception:
                            continue
    except Exception:
        pass
    return cols


# ==========================================
# 3. 버튼용 벡터 라인 아이콘 생성기 (타이트 크롭)
# ==========================================
def create_reload_icon(size=(13, 13), color="#4E5968"):
    """미니멀 회전 화살표 아이콘 (여백 타이트 크롭)"""
    scale = 4
    s = size[0] * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    w = int(1.8 * scale)

    draw.arc([1 * scale, 1 * scale, s - 1 * scale, s - 1 * scale], start=20, end=305, fill=color, width=w)
    tip_x = s - 1 * scale
    tip_y = int(4 * scale)
    barb = int(3.5 * scale)
    draw.line([(tip_x - barb, tip_y - int(1 * scale)), (tip_x, tip_y)], fill=color, width=w)
    draw.line([(tip_x - int(1 * scale), tip_y + barb), (tip_x, tip_y)], fill=color, width=w)

    bbox = img.getbbox()
    cropped = img.crop(bbox) if bbox else img
    res = cropped.resize(size, Image.Resampling.LANCZOS)
    return ctk.CTkImage(light_image=res, dark_image=res, size=size)


def create_folder_icon(size=(13, 13), color="#333D4B"):
    """미니멀 오픈 폴더 아이콘 (여백 타이트 크롭)"""
    scale = 4
    s = size[0] * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    w = int(1.8 * scale)

    back_rect = [int(3.5 * scale), int(1 * scale), int(12.5 * scale), int(13 * scale)]
    draw.rounded_rectangle(back_rect, radius=int(2 * scale), outline=color, width=w)

    front_rect = [int(1 * scale), int(1 * scale), int(8.5 * scale), int(13 * scale)]
    draw.rounded_rectangle(front_rect, radius=int(2 * scale), fill=(242, 244, 246, 255), outline=color, width=w)

    bbox = img.getbbox()
    cropped = img.crop(bbox) if bbox else img
    res = cropped.resize(size, Image.Resampling.LANCZOS)
    return ctk.CTkImage(light_image=res, dark_image=res, size=size)


# ==========================================
# 4. 완벽한 좌우 대칭 아래 화살표(⬇) OptionMenu
# ==========================================
class ModernOptionMenu(ctk.CTkOptionMenu):
    """정렬 방식 선택 드롭다운 - 완벽한 좌우 대칭 아래 화살표(⬇) 렌더링"""
    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates)
        try:
            if self._canvas.find_withtag("dropdown_arrow"):
                self._canvas.itemconfigure("dropdown_arrow", state="hidden")

            w = int(self._current_width)
            h = int(self._current_height)
            cx = w - 18
            cy = h // 2

            self._canvas.delete("custom_arrow")

            # 1. 세로 기둥선
            self._canvas.create_line(
                cx, cy - 5, cx, cy + 4,
                fill="#4E5968", width=2, capstyle="round",
                tags="custom_arrow"
            )
            # 2. 좌측 날개 (독립 대칭 선분)
            self._canvas.create_line(
                cx - 4, cy, cx, cy + 4,
                fill="#4E5968", width=2, capstyle="round",
                tags="custom_arrow"
            )
            # 3. 우측 날개 (좌측과 완벽한 정수 대칭)
            self._canvas.create_line(
                cx + 4, cy, cx, cy + 4,
                fill="#4E5968", width=2, capstyle="round",
                tags="custom_arrow"
            )

            self._canvas.tag_raise("custom_arrow")

            if hasattr(self, "_clicked"):
                self._canvas.tag_bind("custom_arrow", "<Button-1>", self._clicked)
        except Exception:
            pass


# ==========================================
# 5. 툴팁 (Tooltip) 헬퍼 클래스
# ==========================================
class ToolTip:
    """원형 물음표(?) 호버 시 한 줄로 출력되는 카드형 툴팁"""
    def __init__(self, widget, text, font_family):
        self.widget = widget
        self.text = text
        self.font_family = font_family
        self.tip_window = None

        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)
        if hasattr(self.widget, "_label"):
            self.widget._label.bind("<Enter>", self.show_tip)
            self.widget._label.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() - 6
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)

        frame = tk.Frame(tw, bg="#191F28", padx=8, pady=4)
        frame.pack()

        label = tk.Label(
            frame,
            text=self.text,
            font=(self.font_family, 9),
            fg="#FFFFFF",
            bg="#191F28",
            justify="left"
        )
        label.pack()

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


# ==========================================
# 6. 디자인 가이드 적용 모달 팝업들
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
    """안내/경고/오류 통일 모달 팝업"""
    def __init__(self, parent, message, app_font, title_text="안내"):
        super().__init__(parent)

        self.title("안내")
        self.geometry("420x240")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 420) // 2}+{py + (ph - 240) // 2}")

        self.header_label = ctk.CTkLabel(
            self,
            text=title_text,
            font=ctk.CTkFont(family=app_font, size=18, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(22, 8))

        justify_mode = "left" if ("-" in message or "•" in message) else "center"

        self.body_label = ctk.CTkLabel(
            self,
            text=message,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#4E5968",
            justify=justify_mode,
            wraplength=360
        )
        self.body_label.pack(expand=True, padx=24, pady=(0, 14))

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


class SortConflictDialog(ctk.CTkToplevel):
    """정렬 열을 입력했으나 '정렬 안 함'인 경우 확인 팝업"""
    def __init__(self, parent, col_name, app_font):
        super().__init__(parent)
        self.result = False

        self.title("정렬 설정 확인")
        self.geometry("420x240")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 420) // 2}+{py + (ph - 240) // 2}")

        self.header_label = ctk.CTkLabel(
            self,
            text="정렬 설정 확인",
            font=ctk.CTkFont(family=app_font, size=19, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(24, 8))

        display_col = col_name.strip().upper()
        if not display_col.endswith("열"):
            display_col += "열"

        body_text = f"“{display_col}을 입력하셨으나 정렬 안함으로 선택 하였습니다.\n이대로 진행 하시겠습니까?”"
        self.body_label = ctk.CTkLabel(
            self,
            text=body_text,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#4E5968",
            justify="center",
            wraplength=360
        )
        self.body_label.pack(expand=True, padx=20, pady=(0, 18))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 22))

        self.btn_select = ctk.CTkButton(
            btn_frame,
            text="정렬 선택",
            width=175,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=12,
            command=self._on_select
        )
        self.btn_select.pack(side="left", padx=(0, 8))

        self.btn_proceed = ctk.CTkButton(
            btn_frame,
            text="그냥 진행",
            width=175,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            text_color="#FFFFFF",
            corner_radius=12,
            command=self._on_proceed
        )
        self.btn_proceed.pack(side="right")

    def _on_select(self):
        self.result = False
        self.destroy()

    def _on_proceed(self):
        self.result = True
        self.destroy()


class SortNoColDialog(ctk.CTkToplevel):
    """정렬 방식(오름차순/내림차순)은 선택했으나 기준 열을 지정하지 않은 경우 확인 팝업"""
    def __init__(self, parent, app_font):
        super().__init__(parent)
        self.result = False

        self.title("정렬 설정 확인")
        self.geometry("420x240")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        self.geometry(f"+{px + (pw - 420) // 2}+{py + (ph - 240) // 2}")

        self.header_label = ctk.CTkLabel(
            self,
            text="정렬 열 미선택",
            font=ctk.CTkFont(family=app_font, size=19, weight="bold"),
            text_color="#191F28"
        )
        self.header_label.pack(pady=(24, 8))

        body_text = "데이터 정렬 옵션에서 열을 선택하지 않았습니다.\n그냥 진행할까요?"
        self.body_label = ctk.CTkLabel(
            self,
            text=body_text,
            font=ctk.CTkFont(family=app_font, size=12),
            text_color="#4E5968",
            justify="center",
            wraplength=360
        )
        self.body_label.pack(expand=True, padx=20, pady=(0, 18))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 22))

        self.btn_select = ctk.CTkButton(
            btn_frame,
            text="열 선택",
            width=175,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=12,
            command=self._on_select
        )
        self.btn_select.pack(side="left", padx=(0, 8))

        self.btn_proceed = ctk.CTkButton(
            btn_frame,
            text="그냥 진행",
            width=175,
            height=44,
            font=ctk.CTkFont(family=app_font, size=13, weight="bold"),
            fg_color="#3182F6",
            hover_color="#1B64DA",
            text_color="#FFFFFF",
            corner_radius=12,
            command=self._on_proceed
        )
        self.btn_proceed.pack(side="right")

    def _on_select(self):
        self.result = False
        self.destroy()

    def _on_proceed(self):
        self.result = True
        self.destroy()


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
    """작업 완료 안내 모달 팝업 (본문 항목 좌측 정렬)"""
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
# 7. 메인 애플리케이션
# ==========================================
class ReportMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("엑셀 문서 자동 취합 프로그램")

        # 텍스트 깨짐 방지: 기본 너비 560, 최소 너비 540 고정 (최대 너비는 제한 없음)
        self.geometry("560 x 830")
        self.minsize(540, 720)
        self.resizable(True, True)
        self.configure(fg_color="#F2F4F6")

        self.selected_files = []
        self.cached_columns = []

        self.save_dir_path = str(Path.home() / "Desktop")
        if not os.path.exists(self.save_dir_path):
            self.save_dir_path = str(Path.home())

        self.icon_reload = create_reload_icon(size=(13, 13), color="#4E5968")
        self.icon_folder = create_folder_icon(size=(13, 13), color="#333D4B")

        self._init_ui()
        self._setup_drag_and_drop()

        # 창 크기 변경 감지 바인딩 (면책 문구 반응형 줄바꿈)
        self.bind("<Configure>", self._on_window_resize)

    def _init_ui(self):
        # 1. 상단 타이틀 카드
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

        # 파일 선택 및 목록 비우기 버튼
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
            image=self.icon_reload,
            compound="left",
            font=ctk.CTkFont(family=APP_FONT, size=12),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#4E5968",
            corner_radius=10,
            width=94,
            height=36,
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
            font=(APP_FONT, 10)
        )
        self.file_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # 박스 정중앙 안내 문구 (마침표 1개로 정돈 및 9pt 최적화)
        placeholder_text = "위의 추가 버튼을 클릭하거나, 취합할 파일을 이곳으로 드래그 앤 드롭해 주세요."
        self.placeholder_label = tk.Label(
            self.file_listbox,
            text=placeholder_text,
            font=(APP_FONT, 9),
            fg="#8B95A1",
            bg="#F9FAFB",
            cursor="hand2"
        )
        self.placeholder_label.place(relx=0.5, rely=0.5, anchor="center")
        self.placeholder_label.bind("<Button-1>", lambda e: self.select_files())

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

        # 저장 위치 (호버 툴팁 물음표 아이콘)
        self.path_header_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.path_header_frame.pack(fill="x", padx=20, pady=(0, 4))

        self.path_title_label = ctk.CTkLabel(
            self.path_header_frame,
            text="저장 위치",
            font=ctk.CTkFont(family=APP_FONT, size=12, weight="bold"),
            text_color="#333D4B"
        )
        self.path_title_label.pack(side="left")

        self.path_tooltip_btn = ctk.CTkLabel(
            self.path_header_frame,
            text="?",
            width=14,
            height=14,
            corner_radius=7,
            fg_color="#E5E8EB",
            text_color="#6B7280",
            font=ctk.CTkFont(family=APP_FONT, size=9, weight="bold")
        )
        self.path_tooltip_btn.pack(side="left", padx=(5, 0))

        path_tip_text = "기본 값은 바탕화면으로 설정되어 있습니다. 변경을 원하시면 폴더 변경을 눌러 선택해주세요."
        ToolTip(self.path_tooltip_btn, path_tip_text, APP_FONT)

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
            image=self.icon_folder,
            compound="left",
            width=88,
            height=36,
            font=ctk.CTkFont(family=APP_FONT, size=12),
            fg_color="#F2F4F6",
            hover_color="#E5E8EB",
            text_color="#333D4B",
            corner_radius=10,
            command=self.change_directory
        )
        self.path_btn.pack(side="right")

        # 데이터 정렬 옵션 (핑크색 '(선택)' + 물음표 아이콘)
        self.sort_header_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.sort_header_frame.pack(fill="x", padx=20, pady=(0, 4))

        self.sort_title_label = ctk.CTkLabel(
            self.sort_header_frame,
            text="데이터 정렬 옵션",
            font=ctk.CTkFont(family=APP_FONT, size=12, weight="bold"),
            text_color="#333D4B"
        )
        self.sort_title_label.pack(side="left")

        self.sort_opt_badge = ctk.CTkLabel(
            self.sort_header_frame,
            text="(선택)",
            font=ctk.CTkFont(family=APP_FONT, size=11, weight="bold"),
            text_color="#F04452"
        )
        self.sort_opt_badge.pack(side="left", padx=(5, 0))

        self.sort_tooltip_btn = ctk.CTkLabel(
            self.sort_header_frame,
            text="?",
            width=14,
            height=14,
            corner_radius=7,
            fg_color="#E5E8EB",
            text_color="#6B7280",
            font=ctk.CTkFont(family=APP_FONT, size=9, weight="bold")
        )
        self.sort_tooltip_btn.pack(side="left", padx=(5, 0))

        sort_tip_text = "선택 사항입니다. 미선택 시 별도의 정렬은 진행되지 않습니다."
        ToolTip(self.sort_tooltip_btn, sort_tip_text, APP_FONT)

        self.sort_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.sort_frame.pack(fill="x", padx=20, pady=(0, 2))

        self.sort_col_var = tk.StringVar()
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

        self.sort_col_entry.bind("<KeyRelease>", lambda e: self.after(10, self._on_sort_text_change))
        self.sort_col_entry.bind("<FocusOut>", lambda e: self._on_sort_text_change())
        self.sort_col_var.trace_add("write", self._on_sort_text_change)

        self.sort_order_menu = ModernOptionMenu(
            self.sort_frame,
            values=["정렬 안 함", "오름차순 (1→9, A→Z)", "내림차순 (9→1, Z→A)"],
            font=ctk.CTkFont(family=APP_FONT, size=11),
            dropdown_font=ctk.CTkFont(family=APP_FONT, size=11),
            fg_color="#F2F4F6",
            button_color="#F2F4F6",
            button_hover_color="#E5E8EB",
            text_color="#333D4B",
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color="#191F28",
            corner_radius=10,
            width=165,
            height=36
        )
        self.sort_order_menu.set("정렬 안 함")
        self.sort_order_menu.pack(side="right")

        # 실시간 상태 안내 라벨 (한글 경고: 빨강 / 선택 열 안내: 파랑)
        self.sort_warn_label = ctk.CTkLabel(
            self.content_card,
            text="",
            font=ctk.CTkFont(family=APP_FONT, size=11, weight="bold"),
            text_color="#3182F6"
        )
        self.sort_warn_label.pack(anchor="w", padx=22, pady=(0, 6))

        # 프로그레스 영역
        self.progress_frame = ctk.CTkFrame(self.content_card, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=20, pady=(0, 8))

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
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 14))

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

        # 4. 맨 하단 면책 문구 (반응형 동적 줄바꿈)
        disclaimer_text = (
            "본 프로그램은 업무 지원을 목적으로 제공되는 도구입니다. "
            "사용자는 본 프로그램을 자유롭게 수정·사용·배포할 수 있으며, "
            "프로그램의 사용으로 인해 발생하는 모든 문제 및 손해에 대한 책임은 사용자에게 있습니다."
        )
        self.disclaimer_label = ctk.CTkLabel(
            self,
            text=disclaimer_text,
            font=ctk.CTkFont(family=APP_FONT, size=10),
            text_color="#8B95A1",
            justify="center"
        )
        self.disclaimer_label.pack(fill="x", padx=16, pady=(0, 14))

    def _on_window_resize(self, event):
        """창 크기 조절 시 맨 하단 면책 문구의 너비를 반응형으로 실시간 계산"""
        if event.widget == self:
            current_w = self.winfo_width()
            # 창 양쪽 마진(32px)을 제외한 가용 너비로 실시간 줄바꿈
            new_wraplength = max(320, current_w - 40)
            self.disclaimer_label.configure(wraplength=new_wraplength)

    def _extract_sample_columns(self):
        """추가된 문서들에서 컬럼 헤더 목록을 안전하게 분석하여 캐싱"""
        self.cached_columns = []
        if not self.selected_files:
            return

        for f in self.selected_files:
            cols = extract_columns_from_file(f)
            if cols:
                self.cached_columns = cols
                break

    def _on_sort_text_change(self, *args):
        """정렬 입력값 실시간 감지 -> 한글 경고(빨강) 또는 선택 열 명칭 안내(파랑)"""
        val = self.sort_col_entry.get().strip()
        if not val:
            val = self.sort_col_var.get().strip()

        if not val:
            self.sort_warn_label.configure(text="")
            return

        if has_korean(val):
            self.sort_warn_label.configure(text="영문으로 입력해 주세요", text_color="#FF3B30")
            return

        col_idx = excel_col_to_index(val)
        clean_col = val.upper().replace("열", "")

        if col_idx is not None:
            if self.cached_columns:
                if 0 <= col_idx < len(self.cached_columns):
                    col_name = self.cached_columns[col_idx]
                    self.sort_warn_label.configure(
                        text=f"선택하신 {clean_col}열은 “{col_name}” 입니다.",
                        text_color="#3182F6"
                    )
                else:
                    self.sort_warn_label.configure(
                        text=f"선택하신 {clean_col}열은 문서의 열 범위를 초과했습니다. (전체 {len(self.cached_columns)}개 열)",
                        text_color="#8B95A1"
                    )
            else:
                if not self.selected_files:
                    self.sort_warn_label.configure(
                        text=f"선택하신 열: {clean_col}열 (취합할 파일을 추가하시면 해당 열의 명칭이 표시됩니다)",
                        text_color="#3182F6"
                    )
                else:
                    self.sort_warn_label.configure(
                        text=f"선택하신 열: {clean_col}열",
                        text_color="#3182F6"
                    )
        else:
            self.sort_warn_label.configure(text="영문 열 문자(예: A, B, Q)로 입력해 주세요", text_color="#FF3B30")

    def open_about_dialog(self):
        AboutDialog(self, APP_FONT)

    def _setup_drag_and_drop(self):
        """Tkinter 메인 루프 충돌 방지 안전한 드래그 앤 드롭"""
        try:
            import windnd

            def on_drop(files):
                self.after(10, self._process_dropped_files, files)

            windnd.hook_dropfiles(self.file_listbox, func=on_drop)
            windnd.hook_dropfiles(self.placeholder_label, func=on_drop)
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
        self.cached_columns.clear()
        self._refresh_listbox()
        self._set_progress(0, "대기 중")
        self._on_sort_text_change()

    def _refresh_listbox(self):
        self.file_listbox.delete(0, tk.END)

        if not self.selected_files:
            self.placeholder_label.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.placeholder_label.place_forget()
            for f in self.selected_files:
                ext = os.path.splitext(f)[1].lower()
                icon = "📦" if ext == ".zip" else "📄"
                self.file_listbox.insert(tk.END, f" {icon}  {os.path.basename(f)}")

        self._extract_sample_columns()
        self._on_sort_text_change()

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
            warn_msg = "병합할 파일(.zip 또는 엑셀 문서)을\n먼저 추가해 주세요."
            WarningDialog(self, warn_msg, APP_FONT, title_text="파일 선택 안내")
            return

        sort_col_input = self.sort_col_entry.get().strip()
        sort_order_mode = self.sort_order_menu.get()

        if sort_col_input and has_korean(sort_col_input):
            warn_msg = "데이터 정렬 값을 영문으로 표시해 주세요.\n(A열을 정렬하고 싶으면 A, B열을 정렬하고 싶으면 B 등)"
            WarningDialog(self, warn_msg, APP_FONT, title_text="정렬 입력 안내")
            self.sort_col_entry.focus()
            return

        if sort_col_input and sort_order_mode == "정렬 안 함":
            dlg = SortConflictDialog(self, sort_col_input, APP_FONT)
            self.wait_window(dlg)
            if not dlg.result:
                return
            else:
                sort_col_input = ""

        if not sort_col_input and sort_order_mode != "정렬 안 함":
            dlg = SortNoColDialog(self, APP_FONT)
            self.wait_window(dlg)
            if not dlg.result:
                self.sort_col_entry.focus()
                return
            else:
                sort_order_mode = "정렬 안 함"

        out_name = self.filename_entry.get().strip()
        if not out_name:
            out_name = "통합_보고서.xlsx"
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"

        save_dir = self.path_entry.get().strip()
        if not os.path.exists(save_dir):
            warn_msg = "유효한 저장 경로를 지정해 주세요."
            WarningDialog(self, warn_msg, APP_FONT, title_text="저장 경로 안내")
            return

        out_path = os.path.join(save_dir, out_name)

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
                                warn_msg = f"[{os.path.basename(zip_path)}] 암호 입력을 취소하여 작업을 중단합니다."
                                WarningDialog(self, warn_msg, APP_FONT, title_text="작업 취소")
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
                    warn_msg = "취합할 수 있는 유효한 엑셀 또는 CSV 데이터가 없습니다."
                    WarningDialog(self, warn_msg, APP_FONT, title_text="데이터 없음")
                    return

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

                self._set_progress(85, "데이터 컬럼 매핑 중...")
                merged_df = pd.concat(dataframes, ignore_index=True, sort=False)

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
            self.after(0, lambda: CompleteDialog(self, total_docs, out_path, APP_FONT, sort_summary))

        except Exception as e:
            self.reset_ui()
            self._set_progress(0, "오류 발생")
            err_msg = str(e)

            if "file is not a zip file" in err_msg.lower() or "badzipfile" in err_msg.lower():
                drm_msg = (
                    "파일을 확인 할 수 없습니다. 아래 내용을 확인해 주세요.\n\n"
                    "- DRM 설정 여부.\n"
                    "- 손상된 파일 여부.\n"
                    "- 암호 설정 여부."
                )
                self.after(0, lambda: WarningDialog(self, drm_msg, APP_FONT, title_text="파일 확인 불가"))
            else:
                warn_msg = f"작업 중 오류가 발생했습니다:\n{err_msg}"
                self.after(0, lambda: WarningDialog(self, warn_msg, APP_FONT, title_text="오류 발생"))

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
