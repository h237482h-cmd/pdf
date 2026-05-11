#!/usr/bin/env python3
"""
PDF 公文資訊擷取工具 - 圖形介面版
輸出格式對應「總收案登記簿」欄位
"""

import os
import re
import glob
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pdfplumber
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception:
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        except Exception:
            pass
    return text


def clean(s):
    """移除所有空白（全形、半形、Tab）"""
    if not s:
        return ""
    s = s.replace("\u3000", "")
    s = re.sub(r"[ \t\r]+", "", s)
    s = re.sub(r"\n+", " ", s).strip()
    return s


def extract_agency_and_docnum(text):
    """擷取來文機關及字號，合併為一個欄位"""
    agency = ""
    doc_num = ""
    for p in [r"發\s*文\s*機\s*關[：:]\s*(.+)", r"機\s*關\s*名\s*稱[：:]\s*(.+)"]:
        m = re.search(p, text, re.MULTILINE)
        if m:
            agency = clean(m.group(1).split("\n")[0])
            break
    for p in [r"發\s*文\s*字\s*號[：:]\s*(.+)", r"字\s*號[：:]\s*(.+)"]:
        m = re.search(p, text, re.MULTILINE)
        if m:
            doc_num = clean(m.group(1).split("\n")[0])
            break
    if agency and doc_num:
        return f"{agency} {doc_num}"
    return agency or doc_num


def extract_subject(text):
    """擷取主旨，只取前兩行，去除所有空白"""
    for p in [
        r"主\s*旨[：:]\s*(.*?)(?=\n\s*(?:說明|辦理|正本|副本|發文|附件))",
        r"主\s*旨[：:]\s*(.+)",
    ]:
        m = re.search(p, text, re.DOTALL | re.MULTILINE)
        if m:
            lines = [l for l in m.group(1).splitlines() if l.strip()]
            return clean(" ".join(lines[:2]))
    return ""


HEADERS = ["編號", "收文日期", "來文機關及字號", "主旨", "備註", "附件件數", "磁片件數", "郵資金額", "分案室  簽收日期"]
COL_WIDTHS = [8, 14, 30, 55, 12, 10, 10, 10, 18]


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "工作表1"

    h_font  = Font(name="微軟正黑體", bold=True, color="FFFFFF", size=11)
    h_fill  = PatternFill("solid", start_color="1F4E79")
    center  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    top_w   = Alignment(vertical="top", wrap_text=True)
    thin    = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
    alt_f   = PatternFill("solid", start_color="EBF3FB")
    plain_f = PatternFill("solid", start_color="FFFFFF")
    d_font  = Font(name="微軟正黑體", size=10)

    for ci, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font, cell.fill, cell.alignment, cell.border = h_font, h_fill, center, thin
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 26

    CENTER_COLS = {1, 2, 6, 7, 8, 9}
    for ri, row_data in enumerate(rows, 2):
        fill = alt_f if ri % 2 == 0 else plain_f
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = d_font
            cell.alignment = center if ci in CENTER_COLS else top_w
            cell.fill = fill
            cell.border = thin
        ws.row_dimensions[ri].height = 45

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"
    return wb


def process_pdfs(pdf_files, output_path, log_fn, progress_fn):
    rows = []
    total = len(pdf_files)
    for idx, pdf_path in enumerate(pdf_files, 1):
        filename = os.path.basename(pdf_path)
        log_fn(f"[{idx}/{total}] {filename}")
        progress_fn(idx, total)
        text = extract_text_from_pdf(pdf_path)
        rows.append([
            idx,
            "",
            extract_agency_and_docnum(text),
            extract_subject(text),
            "", "", "", "", "",
        ])
    build_workbook(rows).save(output_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 公文資訊擷取工具")
        self.resizable(False, False)
        self.configure(bg="#F0F4F8")
        self._pdf_files = []
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 640, 540
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        BLUE, LIGHT = "#1F4E79", "#F0F4F8"
        tk.Label(self, text="📄 PDF 公文資訊擷取工具",
                 font=("微軟正黑體", 16, "bold"), bg=BLUE, fg="white", pady=14).pack(fill="x")
        tk.Label(self, text="自動擷取「來文機關及字號」「主旨」→ 輸出總收案登記簿格式",
                 font=("微軟正黑體", 10), bg=BLUE, fg="#AED6F1", pady=4).pack(fill="x")

        main = tk.Frame(self, bg=LIGHT, padx=24, pady=16)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)

        def label(row, text):
            tk.Label(main, text=text, font=("微軟正黑體", 10, "bold"),
                     bg=LIGHT, fg="#2C3E50").grid(row=row, column=0, sticky="w", pady=(0, 4))

        def btn(parent, text, cmd):
            return tk.Button(parent, text=text, command=cmd, bg="#2980B9", fg="white",
                             font=("微軟正黑體", 9), relief="flat", padx=8, pady=4, cursor="hand2")

        # PDF 來源
        label(0, "選擇 PDF 檔案或資料夾")
        src_f = tk.Frame(main, bg=LIGHT)
        src_f.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        src_f.columnconfigure(0, weight=1)
        self.src_var = tk.StringVar()
        tk.Entry(src_f, textvariable=self.src_var, font=("微軟正黑體", 10),
                 relief="solid", bd=1).grid(row=0, column=0, ipady=5, sticky="ew")
        bf = tk.Frame(src_f, bg=LIGHT)
        bf.grid(row=0, column=1, padx=(8, 0))
        btn(bf, "選擇檔案", self._pick_files).pack(side="left", padx=(0, 4))
        btn(bf, "選擇資料夾", self._pick_folder).pack(side="left")

        # 輸出路徑
        label(2, "輸出 Excel 儲存位置")
        out_f = tk.Frame(main, bg=LIGHT)
        out_f.grid(row=3, column=0, sticky="ew", pady=(0, 16))
        out_f.columnconfigure(0, weight=1)
        self.out_var = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Desktop", "總收案登記簿.xlsx"))
        tk.Entry(out_f, textvariable=self.out_var, font=("微軟正黑體", 10),
                 relief="solid", bd=1).grid(row=0, column=0, ipady=5, sticky="ew")
        btn(out_f, "瀏覽", self._pick_output).grid(row=0, column=1, padx=(8, 0))

        self.run_btn = tk.Button(main, text="▶  開始擷取", command=self._run,
                                  bg=BLUE, fg="white", font=("微軟正黑體", 12, "bold"),
                                  relief="flat", pady=10, cursor="hand2")
        self.run_btn.grid(row=4, column=0, sticky="ew", pady=(0, 14))

        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.grid(row=5, column=0, sticky="ew", pady=(0, 6))

        self.status_var = tk.StringVar(value="請選擇 PDF 後按下「開始擷取」")
        tk.Label(main, textvariable=self.status_var, font=("微軟正黑體", 9),
                 bg=LIGHT, fg="#555", anchor="w").grid(row=6, column=0, sticky="w")

        label(7, "處理記錄")
        log_f = tk.Frame(main, bg=LIGHT)
        log_f.grid(row=8, column=0, sticky="nsew")
        main.rowconfigure(8, weight=1)
        self.log = tk.Text(log_f, height=8, font=("Consolas", 9), relief="solid", bd=1,
                           state="disabled", bg="#1A1A2E", fg="#A8D8EA")
        sb = tk.Scrollbar(log_f, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _pick_files(self):
        files = filedialog.askopenfilenames(title="選擇 PDF", filetypes=[("PDF", "*.pdf")])
        if files:
            self._pdf_files = list(files)
            self.src_var.set(f"已選擇 {len(files)} 個檔案")
            self._log(f"已選擇 {len(files)} 個 PDF")

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="選擇資料夾")
        if folder:
            files = list(set(
                glob.glob(os.path.join(folder, "**/*.pdf"), recursive=True) +
                glob.glob(os.path.join(folder, "*.pdf"))
            ))
            self._pdf_files = files
            self.src_var.set(folder)
            self._log(f"找到 {len(files)} 個 PDF：{folder}")

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="儲存 Excel", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")], initialfile="總收案登記簿.xlsx")
        if path:
            self.out_var.set(path)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _run(self):
        if not self._pdf_files:
            messagebox.showwarning("提示", "請先選擇 PDF！"); return
        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning("提示", "請指定輸出路徑！"); return
        self.run_btn.config(state="disabled", text="處理中…")
        self.progress["value"] = 0
        self._log(f"\n{'='*40}\n開始處理 {len(self._pdf_files)} 個檔案…")
        def worker():
            try:
                process_pdfs(self._pdf_files, out,
                             log_fn=lambda m: self.after(0, self._log, m),
                             progress_fn=lambda i, t: self.after(0, self._set_progress, i, t))
                self.after(0, self._done, out)
            except Exception as e:
                self.after(0, self._error, str(e))
        threading.Thread(target=worker, daemon=True).start()

    def _set_progress(self, cur, total):
        pct = int(cur / total * 100)
        self.progress["value"] = pct
        self.status_var.set(f"處理中… {cur}/{total} ({pct}%)")

    def _done(self, out):
        self.run_btn.config(state="normal", text="▶  開始擷取")
        self.progress["value"] = 100
        self.status_var.set(f"✅ 完成！{out}")
        self._log(f"\n✅ 完成！\n   {out}")
        messagebox.showinfo("完成", f"擷取完成！\n\nExcel 已儲存至：\n{out}")

    def _error(self, msg):
        self.run_btn.config(state="normal", text="▶  開始擷取")
        self.status_var.set("❌ 發生錯誤")
        self._log(f"❌ 錯誤：{msg}")
        messagebox.showerror("錯誤", msg)


if __name__ == "__main__":
    App().mainloop()
