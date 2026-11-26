import hashlib
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import fitz
import pandas as pd

Z = 1.5  # preview zoom
TRUE = {"1", "true", "yes", "y", "x", "t"}


class App:
    def __init__(self, root):
        self.root = root
        root.title("Form Builder")
        self.path = None
        self.doc = None
        self.page = 0
        self.fields = []  # {page,x,y,key,size}  x,y in PDF points (top-left)
        self.df = None
        self.img = None

        bar = tk.Frame(root)
        bar.pack(fill="x")
        for t, c in [
            ("Open PDF", self.open_pdf), ("Load Data", self.load_data),
            ("Save Tpl", self.save_tpl), ("Load Tpl", self.load_tpl),
            ("<", lambda: self.go(-1)), (">", lambda: self.go(1)),
            ("Undo", self.undo), ("Generate", self.generate),
        ]:
            tk.Button(bar, text=t, command=c).pack(side="left")
        self.row = ttk.Combobox(bar, state="readonly", width=40)
        self.row.pack(side="left", padx=4)
        self.status = tk.Label(bar, text="")
        self.status.pack(side="left")

        fr = tk.Frame(root)
        fr.pack(fill="both", expand=True)
        self.cv = tk.Canvas(fr, bg="gray")
        vs = tk.Scrollbar(fr, command=self.cv.yview)
        hs = tk.Scrollbar(fr, orient="horizontal", command=self.cv.xview)
        self.cv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        vs.pack(side="right", fill="y")
        hs.pack(side="bottom", fill="x")
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Button-1>", self.click)
        self.cv.bind("<MouseWheel>", lambda e: self.cv.yview_scroll(-1 * (e.delta // 120), "units"))

    # ---------- load ----------
    def open_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not p:
            return
        self.path = p
        self.doc = fitz.open(p)
        self.page = 0
        self.fields = []
        self.show()

    def load_data(self):
        p = filedialog.askopenfilename(filetypes=[("Data", "*.xlsx *.xls *.csv")])
        if not p:
            return
        self.df = (pd.read_csv(p, dtype=str) if p.lower().endswith(".csv")
                   else pd.read_excel(p, dtype=str)).fillna("")
        vals = ["ALL ROWS"] + [
            f"{i}: {self.df.iloc[i, 0]}" for i in range(len(self.df))]
        self.row["values"] = vals
        self.row.current(1 if len(vals) > 1 else 0)

    # ---------- view ----------
    def show(self):
        if not self.doc:
            return
        pix = self.doc[self.page].get_pixmap(matrix=fitz.Matrix(Z, Z))
        self.img = tk.PhotoImage(data=pix.tobytes("ppm"))
        self.cv.delete("all")
        self.cv.create_image(0, 0, image=self.img, anchor="nw")
        self.cv.configure(scrollregion=(0, 0, pix.width, pix.height))
        for f in self.fields:
            if f["page"] != self.page:
                continue
            x, y = f["x"] * Z, f["y"] * Z
            self.cv.create_rectangle(x, y, x + 60, y + f["size"] * Z, outline="red")
            self.cv.create_text(x + 2, y, text=f["key"], anchor="nw", fill="red",
                                font=("Arial", 8))
        self.status.config(text=f"page {self.page + 1}/{len(self.doc)}  fields {len(self.fields)}")

    def go(self, d):
        if self.doc and 0 <= self.page + d < len(self.doc):
            self.page += d
            self.show()

    # ---------- fields ----------
    def click(self, e):
        if not self.doc:
            return
        x, y = self.cv.canvasx(e.x) / Z, self.cv.canvasy(e.y) / Z
        cols = ", ".join(self.df.columns) if self.df is not None else "(load data first)"
        key = simpledialog.askstring(
            "Field", f"Column name\n  =text for fixed\n  check:Column for checkbox\n\n{cols}")
        if not key:
            return
        size = simpledialog.askinteger("Size", "Font size", initialvalue=10) or 10
        self.fields.append({"page": self.page, "x": x, "y": y, "key": key, "size": size})
        self.show()

    def undo(self):
        if self.fields:
            self.fields.pop()
            self.show()

    def sha(self):
        return hashlib.sha256(open(self.path, "rb").read()).hexdigest()

    def save_tpl(self):
        if not self.doc:
            return
        p = filedialog.asksaveasfilename(defaultextension=".json")
        if p:
            json.dump({"sha256": self.sha(), "fields": self.fields}, open(p, "w"), indent=1)

    def load_tpl(self):
        if not self.doc:
            messagebox.showinfo("", "Open PDF first")
            return
        p = filedialog.askopenfilename(filetypes=[("Template", "*.json")])
        if not p:
            return
        t = json.load(open(p))
        if t.get("sha256") != self.sha() and not messagebox.askyesno(
                "Warning", "PDF differs from template. Load anyway?"):
            return
        self.fields = t["fields"]
        self.show()

    # ---------- output ----------
    def render(self, rec, out):
        doc = fitz.open(self.path)  # original untouched
        for f in self.fields:
            k = f["key"]
            page = doc[f["page"]]
            if k.startswith("check:"):
                v = str(rec.get(k[6:], "")).strip().lower()
                txt = "X" if v in TRUE else ""
            elif k.startswith("="):
                txt = k[1:]
            else:
                txt = str(rec.get(k, ""))
            if txt:
                page.insert_text((f["x"], f["y"] + f["size"]), txt,
                                 fontsize=f["size"], fontname="helv")
        doc.save(out, garbage=3, deflate=True)
        doc.close()

    def generate(self):
        if not self.doc or self.df is None:
            messagebox.showinfo("", "Open PDF and load data first")
            return
        sel = self.row.current()
        rows = range(len(self.df)) if sel <= 0 else [sel - 1]
        d = filedialog.askdirectory(title="Output folder")
        if not d:
            return
        base = os.path.splitext(os.path.basename(self.path))[0]
        for i in rows:
            rec = self.df.iloc[i].to_dict()
            self.render(rec, os.path.join(d, f"{base}_{i + 1}.pdf"))
        messagebox.showinfo("Done", f"{len(rows)} PDF saved.\nPrint at 100% / Actual size.")


if __name__ == "__main__":
    r = tk.Tk()
    r.geometry("1100x800")
    App(r)
    r.mainloop()
