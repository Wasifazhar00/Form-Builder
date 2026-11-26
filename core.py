"""Form Builder core: storage, security, detection, rendering. No GUI."""
import csv
import datetime as dt
import hashlib
import io
import json
import os
import re
import secrets
import shutil
import time
import uuid
import zipfile

import pymupdf as fitz
from cryptography.fernet import Fernet

DATA = os.environ.get("FB_DATA") or os.path.join(os.path.expanduser("~"), "FormBuilderData")
TPL_DIR = os.path.join(DATA, "templates")
OUT_DEFAULT = os.path.join(DATA, "output")
for _d in (DATA, TPL_DIR, OUT_DEFAULT):
    os.makedirs(_d, exist_ok=True)

TYPES = ["text", "date", "checkbox", "radio", "signature", "multiline",
         "dropdown", "currency", "qr", "barcode", "static"]
TRUE = {"1", "true", "yes", "y", "x", "t", "on"}
DATE_IN = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d %b %Y", "%d %B %Y"]
ROLES = ("admin", "editor", "operator")
PERMS = {"template": {"admin", "editor"}, "delete": {"admin", "editor"},
         "users": {"admin"}, "settings": {"admin"}, "audit": {"admin"}, "purge": {"admin"}}


def _p(name):
    return os.path.join(DATA, name)


def safe(s):
    return re.sub(r"[^\w\-]+", "_", str(s)).strip("_")[:40] or "x"


def can(role, action):
    return role in PERMS.get(action, set(ROLES))


# ---------------------------------------------------------------- settings / audit
def get_settings():
    s = {"retention_days": 30, "output_dir": OUT_DEFAULT}
    try:
        with open(_p("settings.json"), encoding="utf-8") as f:
            s.update(json.load(f))
    except Exception:
        pass
    return s


def save_settings(s):
    with open(_p("settings.json"), "w", encoding="utf-8") as f:
        json.dump(s, f, indent=1)


def audit(user, action, detail=""):
    rec = {"t": dt.datetime.now().isoformat(timespec="seconds"), "u": user, "a": action, "d": detail}
    with open(_p("audit.log"), "a", encoding="utf-8") as f:
        f.write(json.dumps(rec) + "\n")


def audit_tail(n=300):
    try:
        with open(_p("audit.log"), encoding="utf-8") as f:
            return f.readlines()[-n:]
    except Exception:
        return []


def purge_old(folder, days):
    n, cutoff = 0, time.time() - days * 86400
    for root, _, files in os.walk(folder):
        for fn in files:
            p = os.path.join(root, fn)
            try:
                if os.path.getmtime(p) < cutoff:
                    os.remove(p)
                    n += 1
            except OSError:
                pass
    return n


# ---------------------------------------------------------------- encryption (PII at rest)
def _dpapi(data, protect):
    import ctypes
    import ctypes.wintypes as wt

    class Blob(ctypes.Structure):
        _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]
    buf = ctypes.create_string_buffer(data, len(data))
    bin_ = Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = Blob()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(bin_), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("DPAPI failed")
    res = ctypes.string_at(out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return res


_fernet = None


def _fer():
    global _fernet
    if _fernet:
        return _fernet
    kp = _p("key.bin")
    if os.path.exists(kp):
        raw = open(kp, "rb").read()
        key = _dpapi(raw[6:], False) if raw.startswith(b"DPAPI:") else raw[4:]
    else:
        key = Fernet.generate_key()
        try:
            blob = b"DPAPI:" + _dpapi(key, True)
        except Exception:
            blob = b"RAW:" + key
        with open(kp, "wb") as f:
            f.write(blob)
    _fernet = Fernet(key)
    return _fernet


def load_data():
    d = {"columns": [], "records": [], "options": {}, "unique": []}
    p = _p("records.enc")
    if os.path.exists(p):
        d.update(json.loads(_fer().decrypt(open(p, "rb").read())))
    if not d["unique"]:
        d["unique"] = [c for c in d["columns"] if re.search(r"(account|a/c).*(no|num)|iban|cnic", c, re.I)]
    return d


def save_data(d):
    tmp = _p("records.tmp")
    with open(tmp, "wb") as f:
        f.write(_fer().encrypt(json.dumps(d).encode("utf-8")))
    os.replace(tmp, _p("records.enc"))


def new_record(d, vals):
    r = {"_id": uuid.uuid4().hex[:10], "_created": time.time()}
    r.update(vals)
    for k in vals:
        if k not in d["columns"]:
            d["columns"].append(k)
    d["records"].append(r)
    return r


def label_of(d, r):
    for c in d["columns"][:4]:
        if str(r.get(c, "")).strip():
            return str(r[c]).strip()
    return r["_id"]


def import_csv(path, d):
    n = 0
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            new_record(d, {k.strip(): (v or "").strip() for k, v in row.items() if k})
            n += 1
    return n


def export_csv(path, d):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(d["columns"])
        for r in d["records"]:
            w.writerow([r.get(c, "") for c in d["columns"]])


def import_xlsx(path, d):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = wb.active.iter_rows(values_only=True)
    head = [str(h).strip() if h is not None else "" for h in next(rows)]
    n = 0
    for row in rows:
        vals = {}
        for h, v in zip(head, row):
            if not h:
                continue
            if v is None:
                v = ""
            elif isinstance(v, (dt.datetime, dt.date)):
                v = v.strftime("%Y-%m-%d")
            elif isinstance(v, float) and v.is_integer():
                v = str(int(v))
            else:
                v = str(v).strip()
            vals[h] = v
        if any(vals.values()):
            new_record(d, vals)
            n += 1
    wb.close()
    return n


def import_table(path, d):
    return import_xlsx(path, d) if path.lower().endswith((".xlsx", ".xlsm")) else import_csv(path, d)


def purge_records(d, days):
    cutoff = time.time() - days * 86400
    keep = [r for r in d["records"] if r.get("_created", time.time()) >= cutoff]
    n = len(d["records"]) - len(keep)
    d["records"] = keep
    return n


# ---------------------------------------------------------------- users
def users():
    try:
        with open(_p("users.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _hash(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, 200_000).hex()


def set_user(name, pw, role):
    u = users()
    salt = secrets.token_bytes(16)
    u[name] = {"salt": salt.hex(), "hash": _hash(pw, salt), "role": role}
    with open(_p("users.json"), "w", encoding="utf-8") as f:
        json.dump(u, f, indent=1)


def del_user(name):
    u = users()
    u.pop(name, None)
    with open(_p("users.json"), "w", encoding="utf-8") as f:
        json.dump(u, f, indent=1)


def login(name, pw):
    u = users().get(name)
    if u and secrets.compare_digest(u["hash"], _hash(pw, bytes.fromhex(u["salt"]))):
        return u["role"]
    return None


# ---------------------------------------------------------------- source -> pdf
def image_to_pdf(path):
    pix = fitz.Pixmap(path)
    w, h = pix.width, pix.height
    xr, yr = pix.xres, pix.yres
    if xr and yr and xr >= 100 and yr >= 100:
        pw, ph = w / xr * 72, h / yr * 72
    else:  # unknown dpi: fit A4 keeping aspect, so printing at 100% is sane
        s = min(595 / w, 842 / h) if h >= w else min(842 / w, 595 / h)
        pw, ph = w * s, h * s
    doc = fitz.open()
    page = doc.new_page(width=pw, height=ph)
    page.insert_image(page.rect, filename=path)
    return doc.tobytes()


def to_pdf_bytes(path):
    if path.lower().endswith(".pdf"):
        raw = open(path, "rb").read()
        fitz.open(stream=raw, filetype="pdf").close()  # validate
        return raw
    return image_to_pdf(path)


# ---------------------------------------------------------------- templates
def tpl_dir(name):
    return os.path.join(TPL_DIR, name)


def list_templates():
    return sorted(n for n in os.listdir(TPL_DIR) if os.path.exists(os.path.join(tpl_dir(n), "meta.json")))


def _write_meta(d, meta):
    with open(os.path.join(d, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=1)


def create_template(name, src_path, user=""):
    name = re.sub(r"[^\w\- ]", "", name).strip()
    if not name:
        raise ValueError("Bad name")
    d = tpl_dir(name)
    if os.path.exists(d):
        raise ValueError("Template exists. Pick another name (form changed = new template/version).")
    to_pdf_bytes(src_path)  # validate first
    os.makedirs(os.path.join(d, "versions"))
    ext = os.path.splitext(src_path)[1].lower()
    dst = os.path.join(d, "source" + ext)
    shutil.copy2(src_path, dst)
    sha = hashlib.sha256(open(dst, "rb").read()).hexdigest()
    meta = {"name": name, "src_file": "source" + ext, "sha256": sha, "version": 1,
            "fields": [], "created": time.time()}
    _write_meta(d, meta)
    audit(user, "template_create", name)
    return load_template(name)


def load_template(name):
    d = tpl_dir(name)
    with open(os.path.join(d, "meta.json"), encoding="utf-8") as f:
        meta = json.load(f)
    src = os.path.join(d, meta["src_file"])
    if hashlib.sha256(open(src, "rb").read()).hexdigest() != meta["sha256"]:
        raise ValueError("Source form file was modified outside the app: " + name)
    meta["pdf"] = to_pdf_bytes(src)
    meta["dir"] = d
    return meta


def save_fields(tpl, fields, user=""):
    d = tpl["dir"]
    with open(os.path.join(d, "meta.json"), encoding="utf-8") as f:
        meta = json.load(f)
    if meta["fields"] != fields:
        meta["version"] += 1
        meta["fields"] = fields
        with open(os.path.join(d, "versions", f"v{meta['version']}.json"), "w", encoding="utf-8") as f:
            json.dump({"version": meta["version"], "user": user, "saved": time.time(), "fields": fields}, f)
        _write_meta(d, meta)
        audit(user, "template_save", f"{meta['name']} v{meta['version']}")
    tpl["fields"], tpl["version"] = fields, meta["version"]
    sync_columns(fields)
    return tpl


def export_template(tpl, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"name": tpl["name"], "version": tpl["version"], "source_sha256": tpl["sha256"],
                   "fields": tpl["fields"]}, f, indent=1)


def read_template_json(path):
    """Accepts full export or a bare list of (possibly minimal) fields."""
    with open(path, encoding="utf-8") as f:
        j = json.load(f)
    raw = j["fields"] if isinstance(j, dict) else j
    out = []
    for x in raw:
        base = mk(x.get("page", 0), x["rect"], x["key"], x.get("type", "text"), "user")
        base.update(x)
        out.append(base)
    return out, (j.get("source_sha256", "") if isinstance(j, dict) else "")


def delete_template(name, user=""):
    shutil.rmtree(tpl_dir(name), ignore_errors=True)
    audit(user, "template_delete", name)


def sync_columns(fields):
    d = load_data()
    changed = False
    for f in fields:
        k = f["key"]
        if f["type"] == "static" or not k:
            continue
        if k not in d["columns"]:
            d["columns"].append(k)
            changed = True
        if f["type"] == "dropdown" and f.get("options"):
            opts = [o.strip() for o in f["options"].split(",") if o.strip()]
            if d["options"].get(k) != opts:
                d["options"][k] = opts
                changed = True
    if changed:
        save_data(d)


# ---------------------------------------------------------------- fields
def mk(page, rect, key, typ="text", conf="user", **kw):
    f = {"id": uuid.uuid4().hex[:8], "page": page, "rect": [round(float(x), 1) for x in rect],
         "key": key, "type": typ, "size": 10, "align": "left", "fmt": "", "match": "",
         "mark": "X", "options": "", "cond": "", "required": False, "upper": False,
         "acro": None, "conf": conf}
    f.update(kw)
    return f


def truthy(v):
    return str(v).strip().lower() in TRUE


def visible(f, rec):
    c = (f.get("cond") or "").strip()
    if not c:
        return True
    if c.startswith("!"):
        return not truthy(rec.get(c[1:].strip(), ""))
    if "=" in c:
        k, v = c.split("=", 1)
        return str(rec.get(k.strip(), "")).strip().lower() == v.strip().lower()
    return truthy(rec.get(c, ""))


def parse_date(v):
    v = v.strip()
    if v.lower() == "today":
        return dt.date.today()
    for fmt in DATE_IN:
        try:
            return dt.datetime.strptime(v, fmt).date()
        except ValueError:
            pass
    return None


def text_value(f, rec):
    t = f["type"]
    v = f["key"] if t == "static" else str(rec.get(f["key"], "") or "")
    v = v.strip()
    if t == "date" and v:
        d = parse_date(v)
        v = d.strftime(f.get("fmt") or "%d/%m/%Y") if d else v
    elif t == "currency" and v:
        try:
            v = (f.get("fmt") or "{:,.2f}").format(float(v.replace(",", "")))
        except ValueError:
            pass
    if f.get("upper"):
        v = v.upper()
    return v


def is_on(f, rec):
    v = str(rec.get(f["key"], "")).strip()
    return v.lower() == f["match"].strip().lower() if f.get("match") else truthy(v)


# ---------------------------------------------------------------- rendering
def _font():
    for p in (r"C:\Windows\Fonts\arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        if os.path.exists(p):
            return "fbfont", p, fitz.Font(fontfile=p)
    return "helv", None, fitz.Font("helv")


FNAME, FFILE, FONT = _font()


def make_code(kind, data):
    buf = io.BytesIO()
    if kind == "qr":
        import qrcode
        qrcode.make(data).save(buf, format="PNG")
    else:
        import barcode
        from barcode.writer import ImageWriter
        barcode.get("code128", data, writer=ImageWriter()).write(buf, options={"write_text": False, "quiet_zone": 1})
    return buf.getvalue()


def _mark(page, rect, style):
    r = fitz.Rect(rect)
    w, h = r.width, r.height
    lw = max(0.8, min(w, h) * 0.12)
    if style == "tick":
        pts = [fitz.Point(r.x0 + w * .2, r.y0 + h * .55), fitz.Point(r.x0 + w * .42, r.y0 + h * .8),
               fitz.Point(r.x0 + w * .82, r.y0 + h * .2)]
        page.draw_polyline(pts, color=(0, 0, 0), width=lw)
    else:
        page.draw_line(fitz.Point(r.x0 + w * .2, r.y0 + h * .2), fitz.Point(r.x0 + w * .8, r.y0 + h * .8), color=(0, 0, 0), width=lw)
        page.draw_line(fitz.Point(r.x0 + w * .8, r.y0 + h * .2), fitz.Point(r.x0 + w * .2, r.y0 + h * .8), color=(0, 0, 0), width=lw)


def _overlay(page, f, rec):
    t, r = f["type"], fitz.Rect(f["rect"])
    if t in ("checkbox", "radio"):
        if is_on(f, rec):
            _mark(page, r, f.get("mark", "X"))
        return
    if t == "signature":
        p = str(rec.get(f["key"], "")).strip()
        if p and os.path.isfile(p):
            page.insert_image(r, filename=p, keep_proportion=True)
        return
    if t in ("qr", "barcode"):
        v = str(rec.get(f["key"], "")).strip()
        if v:
            try:
                page.insert_image(r, stream=make_code(t, v), keep_proportion=True)
            except Exception:
                pass
        return
    txt = text_value(f, rec)
    if not txt:
        return
    size = float(f.get("size") or 10)
    al = {"left": 0, "center": 1, "right": 2}.get(f.get("align", "left"), 0)
    if t == "multiline":
        s = size
        while s >= 5:
            if page.insert_textbox(r, txt, fontsize=s, fontname=FNAME, fontfile=FFILE, align=al, color=(0, 0, 0)) >= 0:
                return
            s -= 0.5
        return
    s = size
    while s > 5 and FONT.text_length(txt, fontsize=s) > r.width:
        s -= 0.5
    while len(txt) > 1 and FONT.text_length(txt, fontsize=s) > r.width:
        txt = txt[:-1]  # truncate rule
    tw = FONT.text_length(txt, fontsize=s)
    x = r.x0 if al == 0 else (r.x0 + (r.width - tw) / 2 if al == 1 else r.x1 - tw)
    page.insert_text((x, r.y0 + r.height / 2 + s * 0.35), txt, fontsize=s, fontname=FNAME, fontfile=FFILE, color=(0, 0, 0))


def _fill_widget(w, f, rec):
    try:
        t = w.field_type
        if t in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
            w.field_value = w.on_state() if is_on(f, rec) else "Off"
        elif t == fitz.PDF_WIDGET_TYPE_SIGNATURE:
            return
        else:
            w.field_value = text_value(f, rec)
        w.update()
    except Exception:
        pass


def _render_af(tpl, rec, flatten):
    """acroforge pipeline: build AcroForm fields -> fill -> flatten. Images/AcroForm-source fields via PyMuPDF."""
    import acroforge as af
    from acroforge import FieldSpec, FieldType
    H = [p.rect.height for p in fitz.open(stream=tpl["pdf"], filetype="pdf")]
    specs, vals, rest = [], {}, []
    for f in tpl["fields"]:
        if not visible(f, rec) or f["page"] >= len(H):
            continue
        t = f["type"]
        if f.get("acro") or t in ("signature", "qr", "barcode"):
            rest.append(f)
            continue
        x0, y0, x1, y1 = f["rect"]
        r = (x0, H[f["page"]] - y1, x1, H[f["page"]] - y0)  # PDF coords: origin bottom-left
        name = "f_" + f["id"]
        if t in ("checkbox", "radio"):
            specs.append(FieldSpec(type=FieldType.CHECKBOX, page=f["page"], rect=r, name=name, export_value="Yes"))
            vals[name] = bool(is_on(f, rec))
        else:
            v = text_value(f, rec)
            if v:
                specs.append(FieldSpec(type=FieldType.TEXT, page=f["page"], rect=r, name=name))
                vals[name] = v
    pdf = tpl["pdf"]
    if specs:
        pdf = af.fill(af.build(pdf, specs), vals)
        if flatten:
            pdf = af.flatten(pdf)
    if rest:
        pdf = render(dict(tpl, pdf=pdf, fields=rest), rec)
    return pdf


def render(tpl, rec, engine="mupdf", flatten=True):
    """Return flattened PDF bytes. Original source never touched."""
    if engine == "acroforge":
        return _render_af(tpl, rec, flatten)
    doc = fitz.open(stream=tpl["pdf"], filetype="pdf")
    af = {}
    for f in tpl["fields"]:
        if f.get("acro"):
            af.setdefault(f["acro"], []).append(f)
    if af:
        for page in doc:
            for w in list(page.widgets() or []):
                for f in af.get(w.field_name, []):
                    if f["page"] == page.number and visible(f, rec):
                        if f["type"] == "radio" and f.get("match") and w.on_state() != f["match"]:
                            continue
                        _fill_widget(w, f, rec)
        try:
            doc.bake(annots=False, widgets=True)
        except Exception:
            pass
    for f in tpl["fields"]:
        if not f.get("acro") and visible(f, rec) and f["page"] < len(doc):
            _overlay(doc[f["page"]], f, rec)
    return doc.tobytes(garbage=3, deflate=True)


def page_png(pdf, pno, zoom):
    doc = fitz.open(stream=pdf, filetype="pdf")
    return doc[pno].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False).tobytes("png"), len(doc)


def merge_pdfs(parts):
    out = fitz.open()
    for b in parts:
        out.insert_pdf(fitz.open(stream=b, filetype="pdf"))
    return out.tobytes(garbage=3, deflate=True)


def generate(d, tpls, recs, outdir, merge=False, make_zip=False, engine="mupdf", flatten=True):
    os.makedirs(outdir, exist_ok=True)
    out = []
    for r in recs:
        lab = safe(label_of(d, r))
        parts = []
        for t in tpls:
            b = render(t, r, engine, flatten)
            if merge:
                parts.append(b)
            else:
                p = os.path.join(outdir, f"{safe(t['name'])}_{lab}_{r['_id'][:4]}.pdf")
                open(p, "wb").write(b)
                out.append(p)
        if merge:
            p = os.path.join(outdir, f"{lab}_{r['_id'][:4]}.pdf")
            open(p, "wb").write(merge_pdfs(parts))
            out.append(p)
    z = None
    if make_zip and out:
        z = os.path.join(outdir, "fillrun_" + time.strftime("%Y%m%d_%H%M%S") + ".zip")
        with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in out:
                zf.write(p, os.path.basename(p))
    return out, z


# ---------------------------------------------------------------- validation
def validate(d, tpls, recs):
    msgs = []
    for r in recs:
        lab = label_of(d, r)
        for t in tpls:
            for f in t["fields"]:
                if f["type"] == "static" or not visible(f, r):
                    continue
                v = str(r.get(f["key"], "")).strip()
                if f.get("required") and not v:
                    msgs.append(f"{lab}: '{f['key']}' required ({t['name']})")
                if f["type"] == "date" and v and parse_date(v) is None:
                    msgs.append(f"{lab}: '{f['key']}' bad date '{v}'")
    for c in d["unique"]:
        seen = {}
        for r in d["records"]:
            v = str(r.get(c, "")).strip()
            if v:
                seen.setdefault(v, []).append(label_of(d, r))
        for v, labs in seen.items():
            if len(labs) > 1:
                msgs.append(f"Duplicate {c} '{v}': {', '.join(labs)}")
    return list(dict.fromkeys(msgs))


# ---------------------------------------------------------------- auto detection
LABELS = re.compile(r"\b(name|father|mother|husband|address|city|date|dob|birth|account|a/c|cnic|nic|passport|"
                    r"phone|mobile|cell|email|amount|signature|sign|branch|occupation|income|nationality|"
                    r"gender|title|designation|purpose|iban|card|ref)\b", re.I)
_ocr = None


def snake(s):
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")[:40] or "field"


def guess_type(label):
    l = label.lower()
    if re.search(r"signature|\bsign\b", l):
        return "signature"
    if re.search(r"\bdate\b|dob|birth", l):
        return "date"
    if re.search(r"amount|income|balance|salary", l):
        return "currency"
    return "text"


def text_lines(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t:
                out.append((*l["bbox"], t))
    return out


def ocr_lines(page):
    global _ocr
    import numpy as np
    from rapidocr_onnxruntime import RapidOCR
    if _ocr is None:
        _ocr = RapidOCR()
    pix = page.get_pixmap(dpi=200, alpha=False, colorspace=fitz.csRGB)
    arr = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)
    res, _ = _ocr(arr)
    k = 72 / 200
    out = []
    for box, text, _s in res or []:
        xs, ys = [p[0] for p in box], [p[1] for p in box]
        out.append((min(xs) * k, min(ys) * k, max(xs) * k, max(ys) * k, text.strip()))
    return out


def label_fields(page, pno, lines, conf):
    hl, sq = [], []
    try:
        for dr in page.get_drawings():
            for it in dr["items"]:
                if it[0] == "l":
                    a, b = it[1], it[2]
                    if abs(a.y - b.y) < 1 and abs(a.x - b.x) > 30:
                        hl.append((min(a.x, b.x), max(a.x, b.x), a.y))
                elif it[0] == "re":
                    r = it[1]
                    if 6 <= r.width <= 16 and 6 <= r.height <= 16 and abs(r.width - r.height) < 2:
                        sq.append(r)
                    elif r.height < 1.5 and r.width > 30:
                        hl.append((r.x0, r.x1, r.y0))
    except Exception:
        pass
    out, W = [], page.rect.width
    for s in sq:  # checkbox squares: label = text just right of square
        for (x0, y0, x1, y1, t) in lines:
            if 0 <= x0 - s.x1 <= 14 and y0 - 4 <= (s.y0 + s.y1) / 2 <= y1 + 4:
                out.append(mk(pno, s, snake(t), "checkbox", conf))
                break
    for ln in lines:
        x0, y0, x1, y1, t = ln
        leader = re.search(r"[_.]{3,}\s*$", t)
        label = t.rstrip(":._-… ").strip()
        if not re.search(r"[A-Za-z]", label) or len(label) > 40:
            continue
        if not (t.endswith(":") or leader or LABELS.search(label)):
            continue
        if any(abs(f["rect"][0] - x0) < 2 and abs(f["rect"][1] - y0) < 2 for f in out):
            continue
        typ, key = guess_type(label), snake(label)
        if leader:
            xs = x0 + (x1 - x0) * (leader.start() / max(len(t), 1)) + 2
            out.append(mk(pno, (xs, y0 - 1, x1, y1 + 1), key, typ, conf))
            continue
        xs = x1 + 4
        ul = [u for u in hl if u[1] > xs + 20 and u[0] <= x1 + 40 and y0 - 3 <= u[2] <= y1 + 8]
        if ul:
            u = ul[0]
            out.append(mk(pno, (max(u[0], xs), u[2] - (y1 - y0) - 1, u[1], u[2] - 1), key, typ, conf))
            continue
        nxt = [o[0] for o in lines if o is not ln and o[0] > xs and o[1] < y1 and o[3] > y0]
        xe = min(nxt) - 4 if nxt else W - 36
        if xe - xs >= 40:
            out.append(mk(pno, (xs, y0 - 1, xe, y1 + 1), key, typ, conf))
    return out


def detect(pdf, ocr=True):
    doc = fitz.open(stream=pdf, filetype="pdf")
    out, notes = [], []
    wt = {fitz.PDF_WIDGET_TYPE_CHECKBOX: "checkbox", fitz.PDF_WIDGET_TYPE_RADIOBUTTON: "radio",
          fitz.PDF_WIDGET_TYPE_SIGNATURE: "signature", fitz.PDF_WIDGET_TYPE_COMBOBOX: "dropdown",
          fitz.PDF_WIDGET_TYPE_LISTBOX: "dropdown"}
    for pno, page in enumerate(doc):
        ws = list(page.widgets() or [])
        for w in ws:
            typ = wt.get(w.field_type, "text")
            f = mk(pno, w.rect, snake(w.field_name), typ, "high", acro=w.field_name)
            if typ == "radio":
                f["match"] = str(w.on_state() or "")
            if typ == "dropdown":
                f["options"] = ",".join(str(x if not isinstance(x, (list, tuple)) else x[0]) for x in (w.choice_values or []))
            out.append(f)
        if ws:
            continue
        lines, src = text_lines(page), "text"
        if len(lines) < 3 and ocr:
            try:
                lines, src = ocr_lines(page), "ocr"
            except Exception as e:
                notes.append(f"OCR unavailable: {e}")
        out += label_fields(page, pno, lines, "low" if src == "ocr" else "medium")
    return out, notes
