"""Form Builder - fill non-editable PDF/image forms from an internal customer database."""
import base64
import copy
import os
import re
import sys
import time
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, simpledialog, ttk

import core

COL = {"high": "#16a34a", "medium": "#f59e0b", "low": "#dc2626", "user": "#2563eb"}
FORMS = [("Forms", "*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff"), ("All", "*.*")]


def form_dialog(parent, title, spec, init=None):
    """spec: (key,label,kind,extra) kind: entry|pw|combo|check|file. Returns dict or None."""
    init = init or {}
    top = tk.Toplevel(parent)
    top.title(title)
    if parent.winfo_viewable():
        top.transient(parent)
    try:
        top.grab_set()
    except tk.TclError:
        pass
    vs = {}
    for i, (key, label, kind, extra) in enumerate(spec):
        tk.Label(top, text=label, anchor="w").grid(row=i, column=0, sticky="w", padx=6, pady=2)
        if kind == "check":
            v = tk.BooleanVar(value=bool(init.get(key)))
            w = tk.Checkbutton(top, variable=v)
        else:
            iv = init.get(key)
            v = tk.StringVar(value="" if iv is None else str(iv))
            if kind == "combo":
                w = ttk.Combobox(top, textvariable=v, values=extra or [], width=32)
            else:
                w = tk.Entry(top, textvariable=v, width=34, show="*" if kind == "pw" else "")
        w.grid(row=i, column=1, sticky="we", padx=6, pady=2)
        if i == 0:
            w.focus_set()
        if kind == "file":
            tk.Button(top, text="...", command=lambda v=v: v.set(filedialog.askopenfilename(parent=top) or v.get())
                      ).grid(row=i, column=2)
        vs[key] = v
    res = {}

    def ok(*_):
        res.update({k: v.get() for k, v in vs.items()})
        top.destroy()
    bf = tk.Frame(top)
    bf.grid(row=len(spec), column=0, columnspan=3, pady=8)
    tk.Button(bf, text="OK", width=10, command=ok).pack(side="left", padx=4)
    tk.Button(bf, text="Cancel", width=10, command=top.destroy).pack(side="left")
    top.bind("<Return>", ok)
    top.bind("<Escape>", lambda e: top.destroy())
    parent.wait_window(top)
    return res or None


def do_login(root):
    if not core.users():
        r = form_dialog(root, "First run: create admin", [("u", "Username", "entry", None),
                        ("p", "Password", "pw", None), ("p2", "Repeat password", "pw", None)])
        if not r or not r["u"].strip() or not r["p"] or r["p"] != r["p2"]:
            messagebox.showerror("Form Builder", "Admin not created (empty or passwords differ).")
            return None
        core.set_user(r["u"].strip(), r["p"], "admin")
        core.audit(r["u"], "admin_created")
        return r["u"].strip(), "admin"
    for _ in range(5):
        r = form_dialog(root, "Login", [("u", "Username", "entry", None), ("p", "Password", "pw", None)])
        if not r:
            return None
        role = core.login(r["u"].strip(), r["p"])
        if role:
            core.audit(r["u"], "login")
            return r["u"].strip(), role
        core.audit(r["u"], "login_failed")
        messagebox.showerror("Login", "Wrong username or password")
    return None


class App:
    def __init__(self, root, user, role):
        self.root, self.user, self.role = root, user, role
        self.can_edit = core.can(role, "template")
        self.data = core.load_data()
        self.tpl, self.fields, self.page, self.zoom = None, [], 0, 1.3
        self.sel, self.drag, self.dirty, self.img, self.npages = None, None, False, None, 1
        self.queue, self.pv_pdf, self.pv_page, self.pv_img = [], None, 0, None
        self.snap = tk.BooleanVar(value=True)
        self.merge, self.zip = tk.BooleanVar(value=True), tk.BooleanVar(value=False)
        self.engine, self.keepfill = tk.StringVar(value="mupdf (exact fit)"), tk.BooleanVar(value=False)
        self.fsearch = tk.StringVar()
        self.search = tk.StringVar()
        root.report_callback_exception = self.on_error
        root.protocol("WM_DELETE_WINDOW", self.quit)

        m = tk.Menu(root)
        am = tk.Menu(m, tearoff=0)
        adm = role == "admin"
        am.add_command(label="Settings / retention", command=self.settings_dlg, state="normal" if adm else "disabled")
        am.add_command(label="Users", command=self.users_dlg, state="normal" if adm else "disabled")
        am.add_command(label="Audit log", command=self.audit_dlg, state="normal" if adm else "disabled")
        am.add_separator()
        am.add_command(label="Exit", command=self.quit)
        m.add_cascade(label=f"{user} ({role})", menu=am)
        root.config(menu=m)

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True)
        self.t1, self.t2, self.t3 = tk.Frame(nb), tk.Frame(nb), tk.Frame(nb)
        nb.add(self.t1, text=" 1. Form templates ")
        nb.add(self.t2, text=" 2. Customer records ")
        nb.add(self.t3, text=" 3. Fill & print ")
        self.build_editor()
        self.build_records()
        self.build_fill()
        self.refresh_records()
        self.refresh_tpl_lists()
        st = core.get_settings()
        n = core.purge_old(st["output_dir"], int(st["retention_days"]))
        if n:
            core.audit(user, "retention_purge", f"{n} files")

    # ------------------------------------------------------------ misc
    def on_error(self, exc, val, tb):
        core.audit(self.user, "error", repr(val)[:200])
        try:
            traceback.print_exception(exc, val, tb)
        except Exception:
            pass
        messagebox.showerror("Error", f"{val}")

    def quit(self):
        if self.dirty and not messagebox.askyesno("Unsaved", "Discard unsaved template changes?"):
            return
        self.root.destroy()

    def show_text(self, title, text):
        t = tk.Toplevel(self.root)
        t.title(title)
        w = tk.Text(t, width=100, height=30)
        w.pack(fill="both", expand=True)
        w.insert("1.0", text)

    # ------------------------------------------------------------ template editor
    def build_editor(self):
        f = self.t1
        bar = tk.Frame(f)
        bar.pack(fill="x")
        st = "normal" if self.can_edit else "disabled"
        for t, c in [("New from file", self.new_file), ("New from camera", self.new_cam),
                     ("Auto-detect", self.autodetect), ("Save", self.save_tpl),
                     ("Delete field", self.del_field), ("Delete template", self.del_tpl),
                     ("Export JSON", self.exp_tpl), ("Import JSON", self.imp_tpl)]:
            tk.Button(bar, text=t, command=c, state=st).pack(side="left")
        self.tcombo = ttk.Combobox(bar, state="readonly", width=24)
        self.tcombo.pack(side="left", padx=6)
        self.tcombo.bind("<<ComboboxSelected>>", self.open_tpl)
        for t, c in [("<", lambda: self.pg(-1)), (">", lambda: self.pg(1)), ("-", lambda: self.zm(-.2)), ("+", lambda: self.zm(.2))]:
            tk.Button(bar, text=t, width=2, command=c).pack(side="left")
        tk.Checkbutton(bar, text="Snap", variable=self.snap).pack(side="left")
        self.einfo = tk.Label(bar, text="")
        self.einfo.pack(side="left", padx=6)

        pw = tk.PanedWindow(f, orient="horizontal")
        pw.pack(fill="both", expand=True)
        cf = tk.Frame(pw)
        self.cv = tk.Canvas(cf, bg="#888")
        vs = tk.Scrollbar(cf, command=self.cv.yview)
        hs = tk.Scrollbar(cf, orient="horizontal", command=self.cv.xview)
        self.cv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        vs.pack(side="right", fill="y")
        hs.pack(side="bottom", fill="x")
        self.cv.pack(fill="both", expand=True)
        pw.add(cf, stretch="always")
        side = tk.Frame(pw)
        tk.Label(side, text="Fields (drag on form to add)\ngreen=AcroForm orange=text red=OCR blue=reviewed",
                 justify="left", font=("Arial", 8)).pack(anchor="w")
        self.flist = tk.Listbox(side, width=34, exportselection=False)
        self.flist.pack(fill="both", expand=True)
        self.flist.bind("<<ListboxSelect>>", self.flist_sel)
        self.flist.bind("<Double-Button-1>", lambda e: self.sel and self.props(self.sel))
        pw.add(side)
        c = self.cv
        c.bind("<ButtonPress-1>", self.on_down)
        c.bind("<B1-Motion>", self.on_move)
        c.bind("<ButtonRelease-1>", self.on_up)
        c.bind("<Double-Button-1>", lambda e: self.sel and self.props(self.sel))
        c.bind("<MouseWheel>", lambda e: c.yview_scroll(-1 * (e.delta // 120), "units"))
        c.bind("<Delete>", lambda e: self.del_field())
        for k, dx, dy in [("Left", -1, 0), ("Right", 1, 0), ("Up", 0, -1), ("Down", 0, 1)]:
            c.bind(f"<{k}>", lambda e, dx=dx, dy=dy: self.nudge(dx * .5, dy * .5))
            c.bind(f"<Shift-{k}>", lambda e, dx=dx, dy=dy: self.nudge(dx * 5, dy * 5))

    def exp_tpl(self):
        if self.tpl:
            p = filedialog.asksaveasfilename(defaultextension=".json", initialfile=self.tpl["name"] + ".json")
            if p:
                core.export_template({**self.tpl, "fields": self.fields}, p)

    def imp_tpl(self):
        if not self.tpl:
            messagebox.showinfo("Template", "Open or create a template (load the blank form) first.")
            return
        p = filedialog.askopenfilename(filetypes=[("Template JSON", "*.json")])
        if not p:
            return
        fields, sha = core.read_template_json(p)
        if sha and sha != self.tpl["sha256"] and not messagebox.askyesno(
                "Warning", "JSON was made for a different form file. Load anyway?"):
            return
        self.fields, self.sel, self.dirty = fields, None, True
        self.refresh_flist()
        self.redraw()

    def refresh_tpl_lists(self):
        names = core.list_templates()
        self.tcombo["values"] = names
        self.fl.delete(0, "end")
        for n in names:
            self.fl.insert("end", n)

    def new_file(self):
        p = filedialog.askopenfilename(filetypes=FORMS)
        if p:
            self.new_from(p)

    def new_cam(self):
        try:
            import cv2
        except Exception as e:
            messagebox.showerror("Camera", f"OpenCV missing: {e}")
            return
        cap = cv2.VideoCapture(0)
        win, path = "SPACE = capture, ESC = cancel", None
        while cap.isOpened():
            ok, frame = cap.read()
            if not ok:
                break
            cv2.imshow(win, frame)
            k = cv2.waitKey(1) & 255
            if k == 32:
                d = os.path.join(core.DATA, "captures")
                os.makedirs(d, exist_ok=True)
                path = os.path.join(d, time.strftime("cap_%Y%m%d_%H%M%S.jpg"))
                cv2.imwrite(path, frame)
                break
            if k == 27 or cv2.getWindowProperty(win, cv2.WND_PROP_VISIBLE) < 1:
                break
        cap.release()
        cv2.destroyAllWindows()
        if path:
            self.new_from(path)

    def new_from(self, p):
        name = simpledialog.askstring("Template", "Template name:", initialvalue=os.path.splitext(os.path.basename(p))[0])
        if not name:
            return
        try:
            t = core.create_template(name, p, self.user)
        except Exception as e:
            messagebox.showerror("Template", str(e))
            return
        self.refresh_tpl_lists()
        self.load_tpl(t)

    def open_tpl(self, _=None):
        if self.dirty and not messagebox.askyesno("Unsaved", "Discard unsaved changes?"):
            self.tcombo.set(self.tpl["name"] if self.tpl else "")
            return
        try:
            self.load_tpl(core.load_template(self.tcombo.get()))
        except Exception as e:
            messagebox.showerror("Template", str(e))

    def load_tpl(self, t):
        self.tpl, self.fields = t, copy.deepcopy(t["fields"])
        self.page, self.sel, self.dirty = 0, None, False
        self.tcombo.set(t["name"])
        self.redraw()
        self.refresh_flist()

    def save_tpl(self):
        if not self.tpl:
            return
        core.save_fields(self.tpl, copy.deepcopy(self.fields), self.user)
        self.dirty = False
        self.data = core.load_data()
        self.refresh_records()
        self.redraw()
        messagebox.showinfo("Saved", f"{self.tpl['name']} saved (version {self.tpl['version']}).")

    def del_tpl(self):
        if self.tpl and messagebox.askyesno("Delete", f"Delete template '{self.tpl['name']}' and its stored form?"):
            core.delete_template(self.tpl["name"], self.user)
            self.tpl, self.fields, self.dirty = None, [], False
            self.tcombo.set("")
            self.cv.delete("all")
            self.refresh_tpl_lists()
            self.refresh_flist()

    def redraw(self):
        self.cv.delete("all")
        if not self.tpl:
            return
        png, self.npages = core.page_png(self.tpl["pdf"], self.page, self.zoom)
        self.img = tk.PhotoImage(data=base64.b64encode(png))
        self.cv.create_image(0, 0, image=self.img, anchor="nw")
        self.cv.configure(scrollregion=(0, 0, self.img.width(), self.img.height()))
        self.einfo.config(text=f"{self.tpl['name']} v{self.tpl['version']}  page {self.page + 1}/{self.npages}")
        self.draw_fields()

    def draw_fields(self):
        self.cv.delete("fld")
        z = self.zoom
        for f in self.fields:
            if f["page"] != self.page:
                continue
            x0, y0, x1, y1 = [v * z for v in f["rect"]]
            c = COL.get(f["conf"], "#2563eb")
            self.cv.create_rectangle(x0, y0, x1, y1, outline=c, width=3 if f is self.sel else 1, tags="fld")
            self.cv.create_text(x0 + 2, y0 - 1, text=f["key"], anchor="sw", fill=c, font=("Arial", 8), tags="fld")

    def refresh_flist(self):
        self.flist.delete(0, "end")
        for f in self.fields:
            self.flist.insert("end", f"p{f['page'] + 1} {f['key']} [{f['type']}]")
        if self.sel in self.fields:
            self.flist.selection_set(self.fields.index(self.sel))

    def flist_sel(self, _):
        s = self.flist.curselection()
        if s:
            self.sel = self.fields[s[0]]
            if self.sel["page"] != self.page:
                self.page = self.sel["page"]
                self.redraw()
            self.draw_fields()

    def pg(self, d):
        if self.tpl and 0 <= self.page + d < self.npages:
            self.page += d
            self.redraw()

    def zm(self, d):
        self.zoom = max(.5, min(3, self.zoom + d))
        self.redraw()

    def pt(self, e):
        return self.cv.canvasx(e.x) / self.zoom, self.cv.canvasy(e.y) / self.zoom

    def sn(self, v):
        return round(v / 5) * 5 if self.snap.get() else v

    def hit(self, x, y):
        for f in reversed(self.fields):
            r = f["rect"]
            if f["page"] == self.page and r[0] - 1 <= x <= r[2] + 1 and r[1] - 1 <= y <= r[3] + 1:
                return f

    def on_down(self, e):
        self.cv.focus_set()
        if not self.tpl or not self.can_edit:
            return
        x, y = self.pt(e)
        f = self.hit(x, y)
        if f:
            self.sel = f
            r = f["rect"]
            corner = abs(x - r[2]) < 6 / self.zoom and abs(y - r[3]) < 6 / self.zoom
            self.drag = ("size" if corner else "move", x, y, list(r))
        else:
            self.sel = None
            self.drag = ("new", x, y, None)
        self.draw_fields()

    def on_move(self, e):
        if not self.drag:
            return
        mode, x0, y0, r = self.drag
        x, y = self.pt(e)
        if mode == "new":
            self.cv.delete("tmp")
            z = self.zoom
            self.cv.create_rectangle(x0 * z, y0 * z, x * z, y * z, outline="#2563eb", dash=(3, 2), tags="tmp")
            return
        f = self.sel
        if mode == "move":
            dx, dy = x - x0, y - y0
            w, h = r[2] - r[0], r[3] - r[1]
            nx, ny = self.sn(r[0] + dx), self.sn(r[1] + dy)
            f["rect"] = [nx, ny, nx + w, ny + h]
        else:
            f["rect"] = [r[0], r[1], max(r[0] + 4, self.sn(x)), max(r[1] + 4, self.sn(y))]
        self.draw_fields()

    def on_up(self, e):
        if not self.drag:
            return
        mode, x0, y0, r = self.drag
        self.drag = None
        self.cv.delete("tmp")
        x, y = self.pt(e)
        if mode == "new":
            a, b = sorted([self.sn(x0), self.sn(x)]), sorted([self.sn(y0), self.sn(y)])
            if a[1] - a[0] >= 6 and b[1] - b[0] >= 4:
                f = core.mk(self.page, (a[0], b[0], a[1], b[1]), "")
                if self.props(f, new=True):
                    self.fields.append(f)
                    self.sel = f
                    self.dirty = True
                    self.refresh_flist()
                    self.draw_fields()
        elif self.sel and self.sel["rect"] != r:
            self.sel["conf"] = "user"
            self.dirty = True
            self.draw_fields()

    def nudge(self, dx, dy):
        if self.sel and self.can_edit:
            r = self.sel["rect"]
            self.sel["rect"] = [r[0] + dx, r[1] + dy, r[2] + dx, r[3] + dy]
            self.sel["conf"] = "user"
            self.dirty = True
            self.draw_fields()

    def del_field(self):
        if self.sel in self.fields and self.can_edit:
            self.fields.remove(self.sel)
            self.sel, self.dirty = None, True
            self.refresh_flist()
            self.draw_fields()

    def props(self, f, new=False):
        if not self.can_edit:
            return False
        spec = [("key", "Data column (or fixed text)", "combo", self.data["columns"]),
                ("type", "Type", "combo", core.TYPES),
                ("size", "Font size", "entry", None),
                ("align", "Align", "combo", ["left", "center", "right"]),
                ("fmt", "Format (date %d/%m/%Y, money {:,.2f})", "entry", None),
                ("match", "Match value (radio/checkbox)", "entry", None),
                ("mark", "Mark style", "combo", ["X", "tick"]),
                ("options", "Dropdown options a,b,c", "entry", None),
                ("cond", "Show only if (Col | !Col | Col=val)", "entry", None),
                ("required", "Required", "check", None),
                ("upper", "UPPERCASE", "check", None)]
        r = form_dialog(self.root, "Field", spec, f)
        if not r or not r["key"].strip():
            return False
        try:
            r["size"] = float(r["size"] or 10)
        except ValueError:
            r["size"] = 10
        f.update(r)
        f["key"] = f["key"].strip()
        f["conf"] = "user"
        self.dirty = True
        self.refresh_flist()
        self.draw_fields()
        return True

    def autodetect(self):
        if not self.tpl:
            return
        self.root.config(cursor="watch")
        self.root.update()
        try:
            new, notes = core.detect(self.tpl["pdf"])
        finally:
            self.root.config(cursor="")

        def ov(a, b):
            w = min(a[2], b[2]) - max(a[0], b[0])
            h = min(a[3], b[3]) - max(a[1], b[1])
            m = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1])) or 1
            return w > 0 and h > 0 and w * h / m > .3
        n = 0
        for f in new:
            if not any(g["page"] == f["page"] and ov(g["rect"], f["rect"]) for g in self.fields):
                self.fields.append(f)
                n += 1
        self.dirty = True
        self.refresh_flist()
        self.draw_fields()
        messagebox.showinfo("Auto-detect", f"{n} suggested fields added.\nGreen = AcroForm (high), orange = text layer (medium), "
                            "red = OCR (low).\nCheck every box: move, resize, double-click to edit, Delete to remove.\n" + "\n".join(notes))

    # ------------------------------------------------------------ records
    def build_records(self):
        f = self.t2
        bar = tk.Frame(f)
        bar.pack(fill="x")
        dl = "normal" if core.can(self.role, "delete") else "disabled"
        pg = "normal" if core.can(self.role, "purge") else "disabled"
        for t, c, s in [("Add", self.rec_add, "normal"), ("Edit", self.rec_edit, "normal"), ("Delete", self.rec_del, dl),
                        ("Columns", self.cols_dlg, dl), ("Unique cols", self.uniq_dlg, dl),
                        ("Import Excel/CSV", self.imp, "normal"), ("Export CSV", self.exp, "normal"),
                        ("Validate all", self.val_all, "normal"), ("Purge old", self.purge_rec, pg)]:
            tk.Button(bar, text=t, command=c, state=s).pack(side="left")
        tk.Label(bar, text="  Search").pack(side="left")
        e = tk.Entry(bar, textvariable=self.search, width=24)
        e.pack(side="left")
        self.search.trace_add("write", lambda *_: self.refresh_records())
        fr = tk.Frame(f)
        fr.pack(fill="both", expand=True)
        self.tv = ttk.Treeview(fr, show="headings", selectmode="extended")
        vs = ttk.Scrollbar(fr, command=self.tv.yview)
        hs = ttk.Scrollbar(fr, orient="horizontal", command=self.tv.xview)
        self.tv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        vs.pack(side="right", fill="y")
        hs.pack(side="bottom", fill="x")
        self.tv.pack(fill="both", expand=True)
        self.tv.bind("<Double-Button-1>", lambda e: self.rec_edit())

    def refresh_records(self):
        cols = self.data["columns"]
        self.tv["columns"] = cols
        for c in cols:
            self.tv.heading(c, text=c)
            self.tv.column(c, width=120, stretch=False)
        self.tv.delete(*self.tv.get_children())
        q = self.search.get().lower()
        for r in self.data["records"]:
            vals = [r.get(c, "") for c in cols]
            if not q or any(q in str(v).lower() for v in vals):
                self.tv.insert("", "end", iid=r["_id"], values=vals)
        self.refresh_fill()

    def rec_by_id(self, i):
        return next((r for r in self.data["records"] if r["_id"] == i), None)

    def rec_form(self, init):
        spec = []
        for c in self.data["columns"]:
            if c in self.data["options"]:
                spec.append((c, c, "combo", self.data["options"][c]))
            elif re.search(r"sig|photo|img|image", c, re.I):
                spec.append((c, c + " (file)", "file", None))
            else:
                spec.append((c, c, "entry", None))
        if not spec:
            messagebox.showinfo("Records", "No columns yet. Use 'Columns' or save a template with fields first.")
            return None
        return form_dialog(self.root, "Record", spec, init)

    def rec_add(self):
        r = self.rec_form({})
        if r:
            n = core.new_record(self.data, r)
            core.save_data(self.data)
            core.audit(self.user, "record_add", n["_id"])
            self.refresh_records()

    def rec_edit(self):
        s = self.tv.selection()
        if not s:
            return
        rec = self.rec_by_id(s[0])
        r = self.rec_form(rec)
        if r:
            rec.update(r)
            core.save_data(self.data)
            core.audit(self.user, "record_edit", rec["_id"])
            self.refresh_records()

    def rec_del(self):
        s = self.tv.selection()
        if s and messagebox.askyesno("Delete", f"Delete {len(s)} record(s)?"):
            self.data["records"] = [r for r in self.data["records"] if r["_id"] not in s]
            core.save_data(self.data)
            core.audit(self.user, "record_delete", ",".join(s))
            self.refresh_records()

    def cols_dlg(self):
        v = simpledialog.askstring("Columns", "Comma separated column names (order = table order):",
                                   initialvalue=", ".join(self.data["columns"]))
        if v is not None:
            self.data["columns"] = [c.strip() for c in v.split(",") if c.strip()]
            core.save_data(self.data)
            self.refresh_records()

    def uniq_dlg(self):
        v = simpledialog.askstring("Unique columns", "Columns that must not repeat (comma separated):",
                                   initialvalue=", ".join(self.data["unique"]))
        if v is not None:
            self.data["unique"] = [c.strip() for c in v.split(",") if c.strip()]
            core.save_data(self.data)

    def imp(self):
        p = filedialog.askopenfilename(filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("All", "*.*")])
        if p:
            n = core.import_table(p, self.data)
            core.save_data(self.data)
            core.audit(self.user, "import_csv", str(n))
            self.refresh_records()

    def exp(self):
        p = filedialog.asksaveasfilename(defaultextension=".csv")
        if p:
            core.export_csv(p, self.data)
            core.audit(self.user, "export_csv", str(len(self.data["records"])))

    def all_tpls(self):
        out = []
        for n in core.list_templates():
            try:
                out.append(core.load_template(n))
            except Exception as e:
                messagebox.showwarning("Template", str(e))
        return out

    def val_all(self):
        m = core.validate(self.data, self.all_tpls(), self.data["records"])
        self.show_text("Validation", "\n".join(m) or "All OK")

    def purge_rec(self):
        d = simpledialog.askinteger("Purge", "Delete records older than N days:", minvalue=1)
        if d and messagebox.askyesno("Purge", f"Delete all records older than {d} days?"):
            n = core.purge_records(self.data, d)
            core.save_data(self.data)
            core.audit(self.user, "record_purge", f"{n} older than {d}d")
            self.refresh_records()

    # ------------------------------------------------------------ fill & print
    def build_fill(self):
        f = self.t3
        pw = tk.PanedWindow(f, orient="horizontal")
        pw.pack(fill="both", expand=True)
        a = tk.Frame(pw)
        tk.Label(a, text="Customers (multi-select = batch)").pack(anchor="w")
        tk.Entry(a, textvariable=self.fsearch).pack(fill="x")
        self.fsearch.trace_add("write", lambda *_: self.refresh_fill())
        self.rl = tk.Listbox(a, selectmode="extended", exportselection=False, width=30)
        self.rl.pack(fill="both", expand=True)
        tk.Button(a, text="Select all customers", command=lambda: self.rl.selection_set(0, "end")).pack(fill="x")
        pw.add(a)
        b = tk.Frame(pw)
        tk.Label(b, text="Forms (multi-select)").pack(anchor="w")
        self.fl = tk.Listbox(b, selectmode="extended", exportselection=False, width=26)
        self.fl.pack(fill="both", expand=True)
        tk.Checkbutton(b, text="One merged PDF per customer", variable=self.merge).pack(anchor="w")
        tk.Checkbutton(b, text="Also make ZIP", variable=self.zip).pack(anchor="w")
        ttk.Combobox(b, textvariable=self.engine, state="readonly",
                     values=["mupdf (exact fit)", "acroforge"]).pack(fill="x")
        tk.Checkbutton(b, text="Keep fillable (acroforge only)", variable=self.keepfill).pack(anchor="w")
        for t, c in [("Preview", self.preview), ("Validate", self.val_sel), ("Generate PDFs", self.generate),
                     ("BATCH: all customers -> ZIP", self.batch_all), ("Print queue", self.print_queue), ("Clear queue", self.clear_queue)]:
            tk.Button(b, text=t, command=c).pack(fill="x")
        tk.Label(b, text="Queue (double-click opens)").pack(anchor="w")
        self.ql = tk.Listbox(b, height=8, exportselection=False)
        self.ql.pack(fill="both", expand=True)
        self.ql.bind("<Double-Button-1>", self.open_q)
        pw.add(b)
        c = tk.Frame(pw)
        bar = tk.Frame(c)
        bar.pack(fill="x")
        tk.Button(bar, text="<", command=lambda: self.pvpg(-1)).pack(side="left")
        tk.Button(bar, text=">", command=lambda: self.pvpg(1)).pack(side="left")
        self.pvl = tk.Label(bar, text="Preview before printing - check alignment")
        self.pvl.pack(side="left")
        self.pvc = tk.Canvas(c, bg="#888")
        self.pvc.pack(fill="both", expand=True)
        pw.add(c, stretch="always")

    def refresh_fill(self):
        q = self.fsearch.get().lower()
        self.rl.delete(0, "end")
        self.rl_ids = []
        for r in self.data["records"]:
            lab = core.label_of(self.data, r)
            if not q or q in lab.lower() or any(q in str(v).lower() for v in r.values()):
                self.rl.insert("end", lab)
                self.rl_ids.append(r["_id"])

    def sel_recs(self):
        return [self.rec_by_id(self.rl_ids[i]) for i in self.rl.curselection()]

    def sel_tpls(self):
        return [core.load_template(self.fl.get(i)) for i in self.fl.curselection()]

    def need(self):
        r, t = self.sel_recs(), self.sel_tpls()
        if not r or not t:
            messagebox.showinfo("Fill", "Select at least one customer and one form.")
            return None
        return r, t

    def preview(self):
        x = self.need()
        if x:
            self.pv_pdf, self.pv_page = core.render(x[1][0], x[0][0], self.eng(), not self.keepfill.get()), 0
            self.show_pv()

    def show_pv(self):
        png, n = core.page_png(self.pv_pdf, self.pv_page, 1.1)
        self.pv_img = tk.PhotoImage(data=base64.b64encode(png))
        self.pvc.delete("all")
        self.pvc.create_image(0, 0, image=self.pv_img, anchor="nw")
        self.pvl.config(text=f"page {self.pv_page + 1}/{n}  (print at 100% / Actual size)")
        self.pv_n = n

    def pvpg(self, d):
        if self.pv_pdf and 0 <= self.pv_page + d < self.pv_n:
            self.pv_page += d
            self.show_pv()

    def val_sel(self):
        x = self.need()
        if x:
            m = core.validate(self.data, x[1], x[0])
            self.show_text("Validation", "\n".join(m) or "All OK")

    def eng(self):
        return "acroforge" if self.engine.get().startswith("acro") else "mupdf"

    def generate(self):
        x = self.need()
        if x:
            self.run(x[0], x[1], self.merge.get(), self.zip.get())

    def batch_all(self):
        t = self.sel_tpls()
        if not t or not self.data["records"]:
            messagebox.showinfo("Batch", "Select form(s) and have records first.")
            return
        if messagebox.askyesno("Batch", f"Fill {len(t)} form(s) for ALL {len(self.data['records'])} customers and make a ZIP?"):
            self.run(list(self.data["records"]), t, False, True)

    def run(self, r, t, merge, zipit):
        m = core.validate(self.data, t, r)
        if m and not messagebox.askyesno("Validation", "\n".join(m[:15]) + "\n\nGenerate anyway?"):
            return
        out = os.path.join(core.get_settings()["output_dir"], time.strftime("%Y%m%d_%H%M%S"))
        self.root.config(cursor="watch")
        self.root.update()
        try:
            paths, z = core.generate(self.data, t, r, out, merge, zipit, self.eng(), not self.keepfill.get())
        finally:
            self.root.config(cursor="")
        self.queue += paths
        self.ql.delete(0, "end")
        for p in self.queue:
            self.ql.insert("end", os.path.basename(p))
        core.audit(self.user, "generate", f"{len(r)} customers x {len(t)} forms ({', '.join(i['name'] for i in t)}) {self.eng()}")
        messagebox.showinfo("Done", f"{len(paths)} PDF(s) in\n{out}" + (f"\nZIP: {z}" if z else "") +
                            "\n\nPrint at 100% / Actual size, not 'Fit to page'.")

    def open_q(self, _):
        s = self.ql.curselection()
        if s and hasattr(os, "startfile"):
            os.startfile(self.queue[s[0]])

    def clear_queue(self):
        self.queue = []
        self.ql.delete(0, "end")

    def print_queue(self):
        if not self.queue:
            return
        if not hasattr(os, "startfile"):
            messagebox.showinfo("Print", "Printing from app works on Windows only. Open the PDFs from the output folder.")
            return
        if not messagebox.askokcancel("Print", f"Print {len(self.queue)} file(s) with the default PDF app.\n"
                                      "Make sure scaling is 'Actual size' / 100%, NOT 'Fit'."):
            return
        for p in self.queue:
            try:
                os.startfile(p, "print")
            except OSError as e:
                messagebox.showerror("Print", f"{os.path.basename(p)}: {e}")
        core.audit(self.user, "print", f"{len(self.queue)} files")

    # ------------------------------------------------------------ admin
    def settings_dlg(self):
        s = core.get_settings()
        r = form_dialog(self.root, "Settings", [("retention_days", "Delete generated PDFs after N days", "entry", None),
                                                ("output_dir", "Output folder", "entry", None)], s)
        if r:
            try:
                s["retention_days"] = max(1, int(r["retention_days"]))
            except ValueError:
                pass
            s["output_dir"] = r["output_dir"] or s["output_dir"]
            core.save_settings(s)
            core.audit(self.user, "settings", "")

    def audit_dlg(self):
        self.show_text("Audit log", "".join(core.audit_tail()))

    def users_dlg(self):
        top = tk.Toplevel(self.root)
        top.title("Users")
        lb = tk.Listbox(top, width=40)
        lb.pack(fill="both", expand=True)

        def fill():
            lb.delete(0, "end")
            for n, u in core.users().items():
                lb.insert("end", f"{n} ({u['role']})")

        def add():
            r = form_dialog(top, "User (existing name = update)", [("u", "Username", "entry", None), ("p", "Password", "pw", None),
                            ("role", "Role", "combo", list(core.ROLES))], {"role": "operator"})
            if r and r["u"].strip() and r["p"] and r["role"] in core.ROLES:
                core.set_user(r["u"].strip(), r["p"], r["role"])
                core.audit(self.user, "user_set", f"{r['u']} {r['role']}")
                fill()

        def rm():
            s = lb.curselection()
            if s:
                n = lb.get(s[0]).split(" (")[0]
                if n == self.user:
                    messagebox.showerror("Users", "Cannot delete yourself.")
                    return
                core.del_user(n)
                core.audit(self.user, "user_delete", n)
                fill()
        tk.Button(top, text="Add / update", command=add).pack(side="left")
        tk.Button(top, text="Delete", command=rm).pack(side="left")
        fill()


def main():
    root = tk.Tk()
    root.withdraw()
    res = do_login(root)
    if not res:
        return
    root.deiconify()
    root.geometry("1380x880")
    root.title("Form Builder")
    App(root, *res)
    root.mainloop()


if __name__ == "__main__":
    main()
