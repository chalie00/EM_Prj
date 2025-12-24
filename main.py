# -*- coding: utf-8 -*-
import os
import random
import re
import threading
import traceback
from copy import copy
from functools import reduce
from math import gcd
from openpyxl.styles import PatternFill
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import cv2
import easyocr
import numpy as np
import openpyxl
import pandas as pd
from tkinter import filedialog, messagebox, ttk
DENOMS = [10000, 5000, 2000, 1000]
TABLE_COLUMNS = ["쿠폰번호", "이름", "만나이", "성별", "10000원권", "5000원권", "2000원권", "1000원권", "남은금액"]
AMOUNT_COLUMNS = ["10000원권", "5000원권", "2000원권", "1000원권"]
AMOUNT_TABLE_COLUMNS = ["No", "제품명", "단가", "수량", "금액"]
@dataclass
class CouponCell:
    image_path: Path
    grid_pos: Tuple[int, int]  # (col, row)
    coupon_number: Optional[str]
    amount: Optional[int]
    raw_texts: List[str]
class OCRProcessor:
    def __init__(self) -> None:
        # easyocr handles Korean + English without needing a local Tesseract binary.
        self.reader = easyocr.Reader(["ko", "en"], gpu=False, verbose=False)
    def split_grid(self, image: np.ndarray, cols: int, rows: int) -> List[Tuple[np.ndarray, Tuple[int, int]]]:
        h, w = image.shape[:2]
        cell_w, cell_h = w // cols, h // rows
        cells = []
        for r in range(rows):
            for c in range(cols):
                x0, y0 = c * cell_w, r * cell_h
                cell = image[y0 : y0 + cell_h, x0 : x0 + cell_w]
                cells.append((cell, (c, r)))
        return cells
    def parse_amount(self, detections: List[Tuple]) -> Optional[int]:
        candidates = []
        for _, text, conf in detections:
            cleaned = text.replace(",", "").replace(" ", "")
            digits = re.sub(r"[^0-9]", "", cleaned)
            if not digits:
                continue
            try:
                value = int(digits)
            except ValueError:
                continue
            candidates.append((value, conf))
        if not candidates:
            return None
        exact_matches = [(value, conf) for value, conf in candidates if value in DENOMS]
        if exact_matches:
            return max(exact_matches, key=lambda x: x[1])[0]
        # choose the candidate that matches closest to allowed denominations
        best = None
        best_score = -1
        for value, conf in candidates:
            closest = min(DENOMS, key=lambda v: abs(v - value))
            score = conf - (abs(closest - value) / 10000)  # small penalty for distance
            if score > best_score:
                best_score = score
                best = closest
        return best
    def parse_coupon_number(self, detections: List[Tuple]) -> Optional[str]:
        label_hits: List[Tuple[str, float]] = []
        numeric_hits: List[Tuple[str, float, float]] = []  # text, conf, vertical_pos
        for bbox, text, conf in detections:
            cleaned = text.strip()
            digits_only = re.sub(r"[^0-9]", "", cleaned)
            if "no" in cleaned.lower():
                if digits_only:
                    label_hits.append((digits_only, conf))
            if digits_only:
                y_positions = [p[1] for p in bbox]
                avg_y = sum(y_positions) / len(y_positions)
                numeric_hits.append((digits_only, conf, avg_y))
        if label_hits:
            return max(label_hits, key=lambda x: x[1])[0] or None
        if numeric_hits:
            # prefer the lowest number on the ticket (likely near the bottom)
            numeric_hits.sort(key=lambda x: (x[2], -x[1]), reverse=True)
            return numeric_hits[0][0] or None
        return None
    def ocr_cell(self, cell_image: np.ndarray) -> Tuple[Optional[int], Optional[str], List[str]]:
        detections = self.reader.readtext(cell_image)
        amount = self.parse_amount(detections)
        number = self.parse_coupon_number(detections)
        if amount is None or amount == 1000:
            amount = self.ocr_amount_focus(cell_image) or amount
        raw = [t for _, t, _ in detections]
        return amount, number, raw
    def ocr_amount_focus(self, cell_image: np.ndarray) -> Optional[int]:
        h, w = cell_image.shape[:2]
        top = cell_image[: int(h * 0.6), :]
        gray = cv2.cvtColor(top, cv2.COLOR_BGR2GRAY)
        resized = cv2.resize(gray, (w * 2, int(h * 0.6) * 2), interpolation=cv2.INTER_CUBIC)
        _, thresh = cv2.threshold(resized, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        detections = self.reader.readtext(thresh, allowlist="0123456789")
        return self.parse_amount(detections)
    def process_image(self, image_path: Path, cols: int, rows: int) -> List[CouponCell]:
        data = np.fromfile(str(image_path), dtype=np.uint8)
        image = cv2.imdecode(data, cv2.IMREAD_COLOR)
        if image is None:
            raise RuntimeError(f"이미지를 불러올 수 없습니다: {image_path}")
        cells = self.split_grid(image, cols, rows)
        results: List[CouponCell] = []
        for cell_img, pos in cells:
            amount, number, raw = self.ocr_cell(cell_img)
            results.append(CouponCell(image_path=image_path, grid_pos=pos, coupon_number=number, amount=amount, raw_texts=raw))
        return results
def normalize_import_df(df: pd.DataFrame) -> pd.DataFrame:
    header_row = None
    for idx, row in df.iterrows():
        if row.astype(str).str.contains("쿠폰번호").any():
            header_row = idx
            break
    if header_row is None:
        raise ValueError("헤더(쿠폰번호)를 찾을 수 없습니다.")
    header = df.iloc[header_row].tolist()
    data = df.iloc[header_row + 1 :].reset_index(drop=True)
    data.columns = header
    for col in TABLE_COLUMNS:
        if col not in data.columns:
            data[col] = None
    data = data[TABLE_COLUMNS]
    # clean coupon number to string without leading/trailing spaces
    data["쿠폰번호"] = data["쿠폰번호"].apply(lambda x: str(x).strip() if pd.notna(x) else x)
    # numeric columns to numbers
    for col in ["10000원권", "5000원권", "2000원권", "1000원권", "남은금액"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0).astype(int)
    return data
def find_header_row(ws: openpyxl.worksheet.worksheet.Worksheet) -> int:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "쿠폰번호":
                return cell.row
    raise ValueError("헤더(쿠폰번호)를 찾을 수 없습니다.")
def build_export_workbook(
    df: pd.DataFrame,
    path: Path,
    template_path: Optional[Path] = None,
    original_df: Optional[pd.DataFrame] = None,
) -> None:
    if template_path and template_path.exists():
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        header_row = find_header_row(ws)
        data_start = header_row + 1
        template_row = data_start
        highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        original_map: Dict[str, Dict[str, int]] = {}
        if original_df is not None and not original_df.empty:
            for _, row in original_df.iterrows():
                coupon_no = str(row.get("쿠폰번호", "")).strip()
                if not coupon_no:
                    continue
                original_map[coupon_no] = {col: int(row.get(col, 0)) for col in AMOUNT_COLUMNS}
        for idx, (_, row) in enumerate(df.iterrows()):
            row_idx = data_start + idx
            if row_idx > ws.max_row:
                ws.append([None] * len(TABLE_COLUMNS))
            for col_idx, col_name in enumerate(TABLE_COLUMNS, 1):
                src_cell = ws.cell(row=template_row, column=col_idx)
                dest_cell = ws.cell(row=row_idx, column=col_idx)
                dest_cell.value = row[col_name]
                if row_idx != template_row:
                    dest_cell._style = copy(src_cell._style)
                if col_name in AMOUNT_COLUMNS:
                    coupon_no = str(row.get("쿠폰번호", "")).strip()
                    original_vals = original_map.get(coupon_no, {})
                    original_value = int(original_vals.get(col_name, 0))
                    new_value = int(row.get(col_name, 0))
                    if new_value != original_value:
                        dest_cell.fill = highlight_fill
        # clear extra old rows
        last_row = data_start + len(df) - 1
        for row_idx in range(last_row + 1, ws.max_row + 1):
            for col_idx in range(1, len(TABLE_COLUMNS) + 1):
                ws.cell(row=row_idx, column=col_idx).value = None
        wb.save(path)
        return
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet()
    empty_row = [None] * len(TABLE_COLUMNS)
    ws.append(empty_row)
    ws.append([None, "참여자 쿠폰 결산 내역"] + [None] * (len(TABLE_COLUMNS) - 2))
    ws.append(empty_row)
    ws.append(empty_row)
    ws.append(TABLE_COLUMNS)
    for _, row in df.iterrows():
        ws.append([row[col] for col in TABLE_COLUMNS])
    wb.save(path)
class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Coupon OCR & Excel Merger")
        try:
            self.root.state("zoomed")  # windows full screen
        except tk.TclError:
            self.root.attributes("-zoomed", True)
        self.root.minsize(1100, 700)
        self.processor = OCRProcessor()
        self.import_df: Optional[pd.DataFrame] = None
        self.ocr_df = pd.DataFrame(columns=TABLE_COLUMNS)
        self.photo_files: List[Path] = []
        self._export_in_progress = False
        self._edit_entry: Optional[tk.Entry] = None
        self._last_export_path: Optional[Path] = None
        self._import_path: Optional[Path] = None
        self.amount_df = pd.DataFrame(columns=AMOUNT_TABLE_COLUMNS)
        self.amount_target_var = tk.StringVar(value="")
        self._load_last_export_path()
        self._build_ui()
        sample_excel = Path("Sample/Sample_Excel.xlsx")
        if sample_excel.exists():
            try:
                self.load_excel(sample_excel)
            except Exception:
                pass
    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        top_bar = ttk.Frame(self.root, padding=10)
        top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.columnconfigure(2, weight=1)

        source_frame = ttk.LabelFrame(top_bar, text="원본/사진 분석", padding=8)
        source_frame.grid(row=0, column=0, padx=(0, 8), sticky="w")

        ttk.Button(source_frame, text="원본 불러오기", command=self.browse_excel).grid(row=0, column=0, padx=4, pady=2)
        ttk.Button(source_frame, text="사진 선택 (복수)", command=self.browse_photos).grid(row=0, column=1, padx=4, pady=2)
        ttk.Label(source_frame, text="Grid:").grid(row=0, column=2, padx=(12, 2), pady=2)
        self.cols_var = tk.IntVar(value=2)
        self.rows_var = tk.IntVar(value=3)
        ttk.Spinbox(source_frame, from_=1, to=5, textvariable=self.cols_var, width=4).grid(row=0, column=3, pady=2)
        ttk.Label(source_frame, text="x").grid(row=0, column=4, pady=2)
        ttk.Spinbox(source_frame, from_=1, to=5, textvariable=self.rows_var, width=4).grid(row=0, column=5, pady=2)
        ttk.Button(source_frame, text="사진 분석", command=self.run_analysis).grid(row=0, column=6, padx=8, pady=2)
        ttk.Button(source_frame, text="내보내기", command=self.export_excel).grid(row=0, column=7, padx=4, pady=2)
        ttk.Button(source_frame, text="폴더 열기", command=self.open_export_folder).grid(row=0, column=8, padx=4, pady=2)

        amount_frame = ttk.LabelFrame(top_bar, text="금액 계산", padding=8)
        amount_frame.grid(row=0, column=1, padx=(0, 8), sticky="w")

        ttk.Button(amount_frame, text="금액 계산 원본 열기", command=self.open_amount_source).grid(row=0, column=0, padx=4, pady=2)
        ttk.Label(amount_frame, text="원하는 금액:").grid(row=0, column=1, padx=(12, 2), pady=2)
        ttk.Entry(amount_frame, textvariable=self.amount_target_var, width=10).grid(row=0, column=2, pady=2)
        ttk.Button(amount_frame, text="금액수정하기", command=self.adjust_amounts).grid(row=0, column=3, padx=8, pady=2)
        ttk.Button(amount_frame, text="금액 내보내기", command=self.export_amount_table).grid(row=0, column=4, padx=4, pady=2)

        self.status_var = tk.StringVar(value="준비 완료")
        ttk.Label(top_bar, textvariable=self.status_var).grid(row=0, column=2, sticky="e")

        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.grid(row=1, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=0)
        main_frame.columnconfigure(2, weight=1)
        main_frame.columnconfigure(3, weight=0)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # Imported excel panel
        self.import_label = ttk.Label(main_frame, text="원본 파일")
        self.import_label.grid(row=0, column=0, columnspan=2, sticky="w")
        self.import_tree, import_vsb = self._make_tree(main_frame)
        self.import_tree.grid(row=1, column=0, sticky="nsew", padx=(0, 2))
        import_vsb.grid(row=1, column=1, sticky="ns")
        self.import_tree.tag_configure("hover", background="#dff0d8")
        self.import_tree.bind("<Motion>", self.on_import_hover)
        self.import_tree.bind("<Leave>", self.on_import_leave)

        ttk.Label(main_frame, text="금액 계산 목록").grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.amount_tree, amount_vsb = self._make_amount_tree(main_frame)
        self.amount_tree.grid(row=3, column=0, sticky="nsew", padx=(0, 2))
        amount_vsb.grid(row=3, column=1, sticky="ns")
        self.amount_tree.bind("<Double-1>", self.on_amount_double_click)
        self._populate_amount_tree()

        # OCR result panel
        ttk.Label(main_frame, text="사진 분석 결과").grid(row=0, column=2, columnspan=2, sticky="w")
        self.ocr_tree, ocr_vsb = self._make_tree(main_frame)
        self.ocr_tree.grid(row=1, column=2, sticky="nsew", padx=(2, 0))
        ocr_vsb.grid(row=1, column=3, sticky="ns")
        self.ocr_tree.bind("<Double-1>", self.on_ocr_double_click)
        self.ocr_tree.bind("<Button-3>", self.on_ocr_right_click)
        self.ocr_menu = tk.Menu(self.root, tearoff=0)
        self.ocr_menu.add_command(label="삭제", command=self.delete_selected_ocr)
        self._populate_tree(self.ocr_tree, self.ocr_df)

        # Missing info panel
        missing_frame = ttk.Frame(self.root, padding=10)
        missing_frame.grid(row=2, column=0, sticky="ew")
        ttk.Label(missing_frame, text="원본 파일에 없는 쿠폰 번호 (이미지 파일명, 위치)").grid(row=0, column=0, sticky="w")
        self.missing_text = tk.Text(missing_frame, height=4)
        self.missing_text.grid(row=1, column=0, sticky="ew")
        missing_frame.columnconfigure(0, weight=1)
    def _make_tree(self, parent: tk.Widget) -> Tuple[ttk.Treeview, ttk.Scrollbar]:
        tree = ttk.Treeview(parent, columns=TABLE_COLUMNS, show="headings", height=12)
        for col in TABLE_COLUMNS:
            tree.heading(col, text=col)
            tree.column(col, width=90, anchor="center")
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        return tree, vsb
    def _make_amount_tree(self, parent: tk.Widget) -> Tuple[ttk.Treeview, ttk.Scrollbar]:
        tree = ttk.Treeview(parent, columns=AMOUNT_TABLE_COLUMNS, show="headings", height=8, selectmode="browse")
        widths = {"No": 25, "제품명": 160, "단가": 80, "수량": 70, "금액": 90}
        for col in AMOUNT_TABLE_COLUMNS:
            tree.heading(col, text=col)
            tree.column(col, width=widths.get(col, 90), anchor="center")
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        return tree, vsb
    def _refresh_amount_numbers(self) -> None:
        if self.amount_df is None:
            return
        self.amount_df = self.amount_df.reset_index(drop=True)
        self.amount_df["No"] = [idx + 1 for idx in range(len(self.amount_df))]
        if "단가" in self.amount_df.columns and "수량" in self.amount_df.columns:
            self.amount_df["금액"] = (self.amount_df["단가"].fillna(0).astype(int)
                                    * self.amount_df["수량"].fillna(0).astype(int))
    def _populate_amount_tree(self) -> None:
        for item in self.amount_tree.get_children():
            self.amount_tree.delete(item)
        if self.amount_df is not None and not self.amount_df.empty:
            self._refresh_amount_numbers()
            for _, row in self.amount_df.iterrows():
                values = [row.get(col, "") for col in AMOUNT_TABLE_COLUMNS]
                self.amount_tree.insert("", "end", values=values)
        total = 0
        if self.amount_df is not None and not self.amount_df.empty:
            total = int(self.amount_df.get("금액", pd.Series([0])).fillna(0).astype(int).sum())
        total_row = [""] * len(AMOUNT_TABLE_COLUMNS)
        total_row[2] = "Total"
        total_row[-1] = str(total)
        self.amount_tree.insert("", "end", values=total_row)
        self.amount_tree.insert("", "end", values=[""] * len(AMOUNT_TABLE_COLUMNS))
    def _populate_tree(self, tree: ttk.Treeview, df: pd.DataFrame) -> None:
        for item in tree.get_children():
            tree.delete(item)
        for _, row in df.iterrows():
            tree.insert("", "end", values=[row.get(col, "") for col in TABLE_COLUMNS])
        if tree is self.ocr_tree:
            tree.insert("", "end", values=[""] * len(TABLE_COLUMNS))
    def browse_photos(self) -> None:
        files = filedialog.askopenfilenames(
            title="쿠폰 사진 선택",
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.tif"), ("All files", "*.*")],
        )
        if files:
            self.photo_files = [Path(f) for f in files]
            self.status_var.set(f"{len(self.photo_files)}장 선택됨")
    def browse_excel(self) -> None:
        file = filedialog.askopenfilename(title="Excel 파일 선택", filetypes=[("Excel", "*.xlsx;*.xls")])
        if file:
            self.load_excel(Path(file))

    def _load_last_export_path(self) -> None:
        path = Path("last_export_path.txt")
        if not path.exists():
            return
        try:
            saved = path.read_text(encoding="utf-8").strip()
        except OSError:
            return
        if saved:
            self._last_export_path = Path(saved)

    def _save_last_export_path(self) -> None:
        if not self._last_export_path:
            return
        try:
            Path("last_export_path.txt").write_text(str(self._last_export_path), encoding="utf-8")
        except OSError:
            pass

    def open_amount_source(self) -> None:
        file = filedialog.askopenfilename(title="금액 계산 원본 선택", filetypes=[("Excel", "*.xlsx;*.xls")])
        if file:
            self.load_amount_excel(Path(file))

    def load_amount_excel(self, path: Path) -> None:
        df = pd.read_excel(path)
        df.columns = [str(col).strip() for col in df.columns]
        if "제품명" not in df.columns or "단가" not in df.columns:
            messagebox.showwarning("경고", "금액 계산 원본에는 '제품명', '단가' 컬럼이 필요합니다.")
            return
        work = df.copy()
        work["제품명"] = work["제품명"].fillna("").astype(str).str.strip()
        work["단가"] = pd.to_numeric(work["단가"], errors="coerce").fillna(0).astype(int)
        if "수량" in work.columns:
            work["수량"] = pd.to_numeric(work["수량"], errors="coerce").fillna(0).astype(int)
        else:
            work["수량"] = 0

        work = work[(work["제품명"] != "") | (work["단가"] > 0) | (work["수량"] > 0)]
        rows: List[Dict[str, object]] = []
        for _, row in work.iterrows():
            rows.append(
                {
                    "No": len(rows) + 1,
                    "제품명": row.get("제품명", ""),
                    "단가": int(row.get("단가", 0)),
                    "수량": int(row.get("수량", 0)),
                    "금액": 0,
                }
            )
        self.amount_df = pd.DataFrame(rows, columns=AMOUNT_TABLE_COLUMNS)
        self._populate_amount_tree()
        self.status_var.set(f"금액 계산 원본 로드: {path.name}")
    def load_excel(self, path: Path) -> None:
        df_raw = pd.read_excel(path, header=None)
        self.import_df = normalize_import_df(df_raw)
        self._import_path = path
        self._populate_tree(self.import_tree, self.import_df)
        self.status_var.set(f"Excel 로드 완료: {path.name}")
    def run_analysis(self) -> None:
        if not self.photo_files:
            messagebox.showwarning("경고", "사진을 먼저 선택하세요.")
            return

        cols, rows = self.cols_var.get(), self.rows_var.get()
        all_cells: List[CouponCell] = []
        try:
            for img_path in self.photo_files:
                cells = self.processor.process_image(img_path, cols, rows)
                all_cells.extend(cells)
        except Exception as exc:
            messagebox.showerror("OCR 오류", str(exc))
            return

        records: List[Dict[str, object]] = []
        for cell in all_cells:
            rec = {col: 0 for col in ["10000원권", "5000원권", "2000원권", "1000원권"]}
            rec.update({"쿠폰번호": cell.coupon_number, "이름": None, "만나이": None, "성별": None, "남은금액": 0})
            if cell.amount in DENOMS:
                rec[f"{cell.amount}원권"] = 1
            rec["이미지"] = cell.image_path.name
            rec["위치"] = f"{cell.grid_pos[0]},{cell.grid_pos[1]}"
            rec["raw"] = " | ".join(cell.raw_texts)
            records.append(rec)

        if not records:
            messagebox.showinfo("안내", "추출된 쿠폰이 없습니다.")
            return

        self.ocr_df = pd.DataFrame(records)
        self.ocr_df = self.ocr_df.reset_index(drop=True)
        # keep display columns
        display_df = self.ocr_df[TABLE_COLUMNS].copy()
        display_df = display_df.fillna("")
        self._populate_tree(self.ocr_tree, display_df)
        self.status_var.set("OCR 완료")
        self._show_missing()
    def on_ocr_double_click(self, event: tk.Event) -> None:
        if self.ocr_df is None:
            return
        row_id = self.ocr_tree.identify_row(event.y)
        col_id = self.ocr_tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_index = int(col_id.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(TABLE_COLUMNS):
            return
        column_name = TABLE_COLUMNS[col_index]
        editable = {"쿠폰번호", "이름", "만나이", "성별", "10000원권", "5000원권", "2000원권", "1000원권"}
        if column_name not in editable:
            return
        bbox = self.ocr_tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.ocr_tree.set(row_id, column_name)
        if column_name in AMOUNT_COLUMNS and value == "0":
            value = ""
        if self._edit_entry is not None:
            self._edit_entry.destroy()
            self._edit_entry = None
        entry = ttk.Entry(self.ocr_tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.focus_set()
        self._edit_entry = entry
        def save_edit(_: Optional[tk.Event] = None) -> None:
            new_value = entry.get().strip()
            if column_name in AMOUNT_COLUMNS:
                if new_value == "":
                    new_value = "0"
                if not new_value.isdigit():
                    messagebox.showerror("오류", "금액 칸은 숫자만 입력하세요.")
                    entry.focus_set()
                    return
            row_index = self.ocr_tree.index(row_id)
            is_new_row = row_index >= len(self.ocr_df)
            if is_new_row:
                new_row = {col: 0 for col in AMOUNT_COLUMNS}
                new_row.update({"쿠폰번호": "", "이름": None, "만나이": None, "성별": None, "남은금액": 0})
                self.ocr_df = pd.concat([self.ocr_df, pd.DataFrame([new_row])], ignore_index=True)
            if column_name == "쿠폰번호":
                self.ocr_df.at[row_index, column_name] = new_value
            elif column_name in AMOUNT_COLUMNS:
                self.ocr_df.at[row_index, column_name] = int(new_value)
            else:
                self.ocr_df.at[row_index, column_name] = new_value
            self.ocr_tree.set(row_id, column_name, new_value)
            if is_new_row:
                self._populate_tree(self.ocr_tree, self.ocr_df)
            entry.destroy()
            self._edit_entry = None
            self._show_missing()
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
    def on_amount_double_click(self, event: tk.Event) -> None:
        row_id = self.amount_tree.identify_row(event.y)
        col_id = self.amount_tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_index = int(col_id.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(AMOUNT_TABLE_COLUMNS):
            return
        column_name = AMOUNT_TABLE_COLUMNS[col_index]
        if column_name in ("No", "금액"):
            return
        bbox = self.amount_tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.amount_tree.set(row_id, column_name)
        if column_name in ("단가", "수량") and value == "0":
            value = ""
        if self._edit_entry is not None:
            self._edit_entry.destroy()
            self._edit_entry = None
        entry = ttk.Entry(self.amount_tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.focus_set()
        self._edit_entry = entry
        def save_edit(_: Optional[tk.Event] = None) -> None:
            new_value = entry.get().strip()
            if column_name in ("단가", "수량"):
                if new_value == "":
                    new_value = "0"
                if not new_value.isdigit():
                    messagebox.showerror("오류", "금액/수량은 숫자로 입력하세요.")
                    entry.focus_set()
                    return
            row_index = self.amount_tree.index(row_id)
            is_new_row = row_index >= len(self.amount_df)
            if is_new_row:
                new_row = {"No": row_index + 1, "제품명": "", "단가": 0, "수량": 0, "금액": 0}
                self.amount_df = pd.concat([self.amount_df, pd.DataFrame([new_row])], ignore_index=True)
            if column_name == "제품명":
                self.amount_df.at[row_index, column_name] = new_value
            else:
                self.amount_df.at[row_index, column_name] = int(new_value)
            self._populate_amount_tree()
            entry.destroy()
            self._edit_entry = None
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
    def on_ocr_right_click(self, event: tk.Event) -> None:
        row_id = self.ocr_tree.identify_row(event.y)
        if row_id:
            self.ocr_tree.selection_set(row_id)
            self.ocr_menu.tk_popup(event.x_root, event.y_root)
    def delete_selected_ocr(self) -> None:
        if self.ocr_df is None:
            return
        selected = self.ocr_tree.selection()
        if not selected:
            return
        row_id = selected[0]
        row_index = self.ocr_tree.index(row_id)
        if row_index >= len(self.ocr_df):
            return
        self.ocr_df = self.ocr_df.drop(self.ocr_df.index[row_index]).reset_index(drop=True)
        self._populate_tree(self.ocr_tree, self.ocr_df)
        self._show_missing()
    def on_import_hover(self, event: tk.Event) -> None:
        row_id = self.import_tree.identify_row(event.y)
        for item in self.import_tree.get_children():
            self.import_tree.item(item, tags=())
        if row_id:
            self.import_tree.item(row_id, tags=("hover",))
    def on_import_leave(self, _: tk.Event) -> None:
        for item in self.import_tree.get_children():
            self.import_tree.item(item, tags=())
    def activate_amount_tab(self) -> None:
        self.import_label.configure(text="금액 계산")
    def _parse_target_amount(self) -> Optional[int]:
        raw = self.amount_target_var.get().strip()
        if raw == "":
            messagebox.showwarning("경고", "원하는 금액을 입력하세요.")
            return None
        if not raw.isdigit():
            messagebox.showwarning("경고", "원하는 금액은 숫자로 입력하세요.")
            return None
        return int(raw)
    def _solve_quantities(self, prices: List[int], target: int, min_each: int = 0) -> Optional[List[int]]:
        if target < 0 or not prices:
            return None
        base_total = sum(price * min_each for price in prices)
        if target < base_total:
            return None
        remaining_target = target - base_total
        denom = reduce(gcd, prices)
        if remaining_target % denom != 0:
            return None
        attempts = 2000
        for _ in range(attempts):
            remaining = remaining_target
            qtys = [min_each] * len(prices)
            for idx in range(len(prices) - 1):
                price = prices[idx]
                max_qty = remaining // price
                qty = random.randint(0, max_qty) if max_qty > 0 else 0
                qtys[idx] += qty
                remaining -= qty * price
            last_price = prices[-1]
            if remaining % last_price == 0:
                qtys[-1] += remaining // last_price
                return qtys
        return None
    def adjust_amounts(self) -> None:
        target = self._parse_target_amount()
        if target is None:
            return
        if self.amount_df is None or self.amount_df.empty:
            messagebox.showinfo("안내", "제품 목록을 먼저 입력하세요.")
            return
        eligible: List[Tuple[int, int]] = []
        for idx, row in self.amount_df.iterrows():
            name = str(row.get("제품명", "")).strip()
            price = int(row.get("단가", 0) or 0)
            if name and price <= 0:
                messagebox.showwarning("경고", "제품명에 해당하는 단가가 0입니다. 단가를 입력하세요.")
                return
            if name and price > 0:
                eligible.append((idx, price))
        if not eligible:
            messagebox.showinfo("안내", "제품명과 금액이 입력된 항목이 없습니다.")
            return

        selected_indices = [idx for idx, _ in eligible]
        prices = [price for _, price in eligible]
        qtys = self._solve_quantities(prices, target, min_each=1)
        if qtys is None:
            messagebox.showwarning("경고", "모든 제품 수량을 1개 이상으로 설정할 수 없습니다.")
            return

        self.amount_df["수량"] = 0
        for idx, qty in zip(selected_indices, qtys):
            self.amount_df.at[idx, "수량"] = int(qty)
        self._populate_amount_tree()
    def export_amount_table(self) -> None:
        if self.amount_df is None or self.amount_df.empty:
            messagebox.showinfo("안내", "내보낼 데이터가 없습니다.")
            return
        save_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="금액 내보내기",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not save_path:
            return
        export_df = self.amount_df.copy()
        export_df = export_df[
            (export_df["제품명"].astype(str).str.strip() != "")
            | (export_df["단가"].fillna(0).astype(int) > 0)
            | (export_df["수량"].fillna(0).astype(int) > 0)
        ].reset_index(drop=True)
        export_df["No"] = [idx + 1 for idx in range(len(export_df))]
        export_df["금액"] = (export_df["단가"].fillna(0).astype(int)
                           * export_df["수량"].fillna(0).astype(int))
        total_amount = int(export_df["금액"].fillna(0).astype(int).sum())

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(AMOUNT_TABLE_COLUMNS)
        for _, row in export_df.iterrows():
            ws.append([row.get(col, "") for col in AMOUNT_TABLE_COLUMNS])
        ws.append(["", "", "Total", "", total_amount])

        thin = openpyxl.styles.Side(style="thin", color="000000")
        border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=len(AMOUNT_TABLE_COLUMNS)):
            for cell in row:
                cell.border = border

        wb.save(save_path)
        self._last_export_path = Path(save_path)
        self._save_last_export_path()
        messagebox.showinfo("완료", f"저장 완료: {Path(save_path).name}")
    def open_export_folder(self) -> None:
        if not self._last_export_path:
            messagebox.showinfo("안내", "Export된 파일이 없습니다.")
            return
        folder = self._last_export_path
        if not folder.is_dir():
            folder = folder.parent
        if folder.exists():
            os.startfile(str(folder))
        else:
            messagebox.showinfo("안내", "저장 폴더를 찾을 수 없습니다.")
    def _show_missing(self) -> None:
        self.missing_text.delete("1.0", tk.END)
        if self.import_df is None or self.ocr_df is None:
            return
        base_numbers = set(self.import_df["쿠폰번호"].dropna().astype(str))
        missing_rows = []
        for _, row in self.ocr_df.iterrows():
            num = str(row["쿠폰번호"]) if pd.notna(row["쿠폰번호"]) else None
            if not num or num not in base_numbers:
                missing_rows.append(
                    f"{row.get('이미지','')}, 위치 {row.get('위치','')}, 쿠폰번호: {num or '미인식'}"
                )
        if missing_rows:
            self.missing_text.insert(tk.END, "\n".join(missing_rows))
        else:
            self.missing_text.insert(tk.END, "모든 쿠폰번호가 Import 파일에 존재합니다.")
    def export_excel(self) -> None:
        if self.import_df is None:
            messagebox.showwarning("경고", "먼저 Excel을 Import 하세요.")
            return
        if self.ocr_df is None:
            messagebox.showwarning("경고", "먼저 OCR Analysis를 실행하세요.")
            return
        if self._export_in_progress:
            messagebox.showinfo("안내", "Export가 진행 중입니다. 잠시만 기다려 주세요.")
            return

        save_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not save_path:
            return
        self._export_in_progress = True
        self.status_var.set("Export 준비 중...")
        self.root.update_idletasks()

        def worker() -> None:
            error_msg = None
            try:
                self.root.after(0, lambda: self.status_var.set("Export 계산 중..."))
                merged = self.merge_data()
                target_path = Path(save_path)
                temp_path = target_path.with_suffix(target_path.suffix + ".tmp")
                log_path = Path("export_error.log")
                log_path.write_text(
                    f"start export: {target_path}\n",
                    encoding="utf-8",
                )
                self.root.after(0, lambda: self.status_var.set("Export 저장 중..."))
                build_export_workbook(
                    merged,
                    temp_path,
                    template_path=self._import_path,
                    original_df=self.import_df,
                )
                log_path.write_text(
                    f"saved temp: {temp_path}\n",
                    encoding="utf-8",
                    errors="replace",
                )
                if not temp_path.exists() or temp_path.stat().st_size == 0:
                    raise RuntimeError("임시 파일 생성에 실패했습니다.")
                temp_path.replace(target_path)
                log_path.write_text(
                    "replaced temp -> target\n",
                    encoding="utf-8",
                    errors="replace",
                )
            except PermissionError:
                error_msg = "파일이 열려 있어 저장할 수 없습니다. Excel에서 닫고 다시 시도하세요."
            except Exception as exc:
                log_path.write_text(traceback.format_exc(), encoding="utf-8")
                error_msg = f"Export 실패: {exc} (export_error.log 확인)"

            def on_done() -> None:
                self._export_in_progress = False
                if error_msg:
                    messagebox.showerror("오류", error_msg)
                    self.status_var.set("Export 실패")
                    return
                self._last_export_path = Path(save_path)
                self._save_last_export_path()
                self.status_var.set(f"Export 완료: {Path(save_path).name}")
                messagebox.showinfo("완료", "Export 완료되었습니다.")

            self.root.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()
    def merge_data(self) -> pd.DataFrame:
        base = self.import_df.copy() if self.import_df is not None else pd.DataFrame(columns=TABLE_COLUMNS)
        # prepare OCR summary grouped by coupon number
        ocr_df = self.ocr_df.copy()
        ocr_df["쿠폰번호"] = ocr_df["쿠폰번호"].fillna("").astype(str)
        ocr_meta = (
            ocr_df.groupby("쿠폰번호")[["이름", "만나이", "성별"]]
            .agg(lambda x: next((v for v in x if pd.notna(v) and str(v).strip() != ""), ""))
            .reset_index()
        )
        grouped = (
            ocr_df.groupby("쿠폰번호")[["10000원권", "5000원권", "2000원권", "1000원권"]]
            .sum()
            .reset_index()
        )
        base["쿠폰번호"] = base["쿠폰번호"].fillna("").astype(str)
        merged = base.merge(grouped, on="쿠폰번호", how="outer", suffixes=("_base", "_ocr"))
        merged = merged.merge(ocr_meta, on="쿠폰번호", how="left", suffixes=("_base", "_ocr"))
        def summed(col: str) -> List[int]:
            base_col = f"{col}_base"
            ocr_col = f"{col}_ocr"
            merged[col] = merged.get(base_col, 0).fillna(0).astype(int) + merged.get(ocr_col, 0).fillna(0).astype(int)
            return merged[col]
        for col in ["10000원권", "5000원권", "2000원권", "1000원권"]:
            summed(col)
        for col in ["이름", "만나이", "성별"]:
            base_col = f"{col}_base"
            ocr_col = f"{col}_ocr"
            base_vals = merged.get(base_col, "").fillna("")
            ocr_vals = merged.get(ocr_col, "").fillna("")
            merged[col] = base_vals.where(base_vals.astype(str).str.strip() != "", ocr_vals)
        # clean up columns
        keep_cols = TABLE_COLUMNS
        for col in keep_cols:
            if col not in merged.columns:
                merged[col] = ""
        merged = merged[keep_cols]
        merged[["이름", "만나이", "성별"]] = merged[["이름", "만나이", "성별"]].fillna("")
        return merged
def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()
if __name__ == "__main__":
    main()
