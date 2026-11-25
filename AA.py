import io
import json
import os
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import fitz  # PyMuPDF
import openpyxl
import pytesseract
from docx import Document
from PIL import Image, ImageDraw, ImageFont, ImageSequence, ImageTk

# Peka ut Tesseract på Windows (används automatiskt om det finns installerat)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

STODDA_FORMAT = {".docx", ".pdf", ".tif", ".tiff", ".xlsx", ".txt", ".md"}

# Mönster för valfri automatisk maskering (envägs, vid anonymisering)
REGEX_EPOST = re.compile(r"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b")
REGEX_TELEFON = re.compile(r"(?:\+46[\s-]?|07\d[\s-]?)\d{3}[\s-]?\d{2}[\s-]?\d{2}\b")
REGEX_KOORDINAT = re.compile(r"\b(?:[NE]\s*[:=]?\s*)?(?:6\d{6}|7\d{6})\b")


# ==========================================
# KÄRNLOGIK (VERSALBEVARANDE & FILFORMAT)
# ==========================================

def matcha_versaler(original: str, ersattning: str) -> str:
    if not original or not ersattning:
        return ersattning
    if original.isupper():
        return ersattning.upper()
    if original.islower():
        return ersattning.lower()
    if original[0].isupper():
        return ersattning[0].upper() + ersattning[1:]
    return ersattning


def skapa_regex(mapping: dict[str, str]) -> re.Pattern:
    sorterade = sorted(mapping.keys(), key=len, reverse=True)
    monster = "|".join(re.escape(k) for k in sorterade)
    return re.compile(monster, re.IGNORECASE)


def ersatt_text(
    text: str,
    mapping: dict[str, str],
    maskera_epost: bool = False,
    maskera_telefon: bool = False,
    maskera_koord: bool = False,
) -> str:
    if not text:
        return text

    resultat = text
    if mapping:
        lookup = {k.lower(): v for k, v in mapping.items()}
        regex = skapa_regex(mapping)
        resultat = regex.sub(
            lambda m: matcha_versaler(m.group(0), lookup[m.group(0).lower()]),
            resultat,
        )

    if maskera_epost:
        resultat = REGEX_EPOST.sub("[E-POST]", resultat)
    if maskera_telefon:
        resultat = REGEX_TELEFON.sub("[TELEFON]", resultat)
    if maskera_koord:
        resultat = REGEX_KOORDINAT.sub("[KOORDINAT]", resultat)

    return resultat


# --- 1. WORD (.docx) ---
def bearbeta_docx(
    in_fil: Path,
    ut_fil: Path,
    mapping: dict[str, str],
    rensa_meta: bool,
    extra_flags: dict[str, bool],
) -> list[str]:
    varningar = []
    doc = Document(in_fil)

    def _ersatt(t: str) -> str:
        return ersatt_text(t, mapping, **extra_flags)

    for p in doc.paragraphs:
        if p.text:
            p.text = _ersatt(p.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text:
                        p.text = _ersatt(p.text)

    for section in doc.sections:
        for hf in (section.header, section.footer):
            for p in hf.paragraphs:
                if p.text:
                    p.text = _ersatt(p.text)

    # Textrutor (flytande objekt i Word XML)
    for txbx_t in doc.element.xpath(".//w:txbxContent//w:t"):
        if txbx_t.text:
            txbx_t.text = _ersatt(txbx_t.text)

    # Räkna inbäddade bilder
    antal_bilder = len(doc.inline_shapes)
    if antal_bilder > 0 and rensa_meta:
        varningar.append(f"Innehåller {antal_bilder} inbäddad(e) bild(er) – kontrollera manuellt.")

    if rensa_meta:
        props = doc.core_properties
        props.author = props.last_modified_by = props.title = ""
        props.subject = props.comments = props.keywords = ""

    doc.save(ut_fil)
    return varningar


# --- 2. EXCEL (.xlsx) ---
def bearbeta_xlsx(
    in_fil: Path,
    ut_fil: Path,
    mapping: dict[str, str],
    rensa_meta: bool,
    extra_flags: dict[str, bool],
) -> list[str]:
    wb = openpyxl.load_workbook(in_fil)

    for sheet in wb.worksheets:
        # Byt även fliknamn (max 31 tecken i Excel)
        nytt_fliknamn = ersatt_text(sheet.title, mapping, **extra_flags)[:31]
        if nytt_fliknamn and nytt_fliknamn != sheet.title:
            sheet.title = nytt_fliknamn

        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and not cell.value.startswith("="):
                    cell.value = ersatt_text(cell.value, mapping, **extra_flags)

    if rensa_meta:
        wb.properties.creator = ""
        wb.properties.lastModifiedBy = ""
        wb.properties.title = ""
        wb.properties.subject = ""
        wb.properties.keywords = ""

    wb.save(ut_fil)
    return []


# --- 3. BILD-HANTERING (TIF och skannade PDF-sidor) ---
def bearbeta_pil_bild(
    sida_rgb: Image.Image,
    mapping: dict[str, str],
    maskera_ritningshuvud: bool,
    huvud_bredd_pct: float,
    huvud_hojd_pct: float,
    extra_flags: dict[str, bool],
) -> Image.Image:
    draw = ImageDraw.Draw(sida_rgb)

    if maskera_ritningshuvud:
        bredd, hojd = sida_rgb.size
        x0 = int(bredd * (1.0 - huvud_bredd_pct / 100.0))
        y0 = int(hojd * (1.0 - huvud_hojd_pct / 100.0))
        draw.rectangle([x0, y0, bredd, hojd], fill="white")

    # Kör bara OCR om Tesseract faktiskt finns installerat på datorn
    if Path(pytesseract.pytesseract.tesseract_cmd).exists():
        data = pytesseract.image_to_data(
            sida_rgb, lang="swe+eng", output_type=pytesseract.Output.DICT
        )

        for i, ord_text in enumerate(data["text"]):
            if not ord_text.strip():
                continue
            nytt_ord = ersatt_text(ord_text, mapping, **extra_flags)
            if nytt_ord != ord_text:
                x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
                draw.rectangle([x - 2, y - 2, x + w + 2, y + h + 2], fill="white")
                try:
                    font = ImageFont.truetype("arial.ttf", max(10, int(h * 0.85)))
                except OSError:
                    font = ImageFont.load_default()
                draw.text((x, y), nytt_ord, fill="black", font=font)

    return sida_rgb


def bearbeta_tif(
    in_fil: Path,
    ut_fil: Path,
    mapping: dict[str, str],
    maskera_ritningshuvud: bool,
    huvud_bredd_pct: float,
    huvud_hojd_pct: float,
    extra_flags: dict[str, bool],
) -> list[str]:
    img = Image.open(in_fil)
    bearbetade_sidor = []

    for sida in ImageSequence.Iterator(img):
        sida_rgb = bearbeta_pil_bild(
            sida.convert("RGB"),
            mapping,
            maskera_ritningshuvud,
            huvud_bredd_pct,
            huvud_hojd_pct,
            extra_flags,
        )
        bearbetade_sidor.append(sida_rgb)

    bearbetade_sidor[0].save(
        ut_fil,
        save_all=True,
        append_images=bearbetade_sidor[1:],
        compression="tiff_lzw",
    )
    return []


# --- 4. PDF (.pdf) med automatisk OCR-fallback ---
def bearbeta_pdf(
    in_fil: Path,
    ut_fil: Path,
    mapping: dict[str, str],
    ocr_fallback: bool,
    maskera_ritningshuvud: bool,
    huvud_bredd_pct: float,
    huvud_hojd_pct: float,
    extra_flags: dict[str, bool],
) -> list[str]:
    varningar = []
    doc = fitz.open(in_fil)
    nytt_doc = fitz.open()
    sorterade_par = sorted(mapping.items(), key=lambda x: len(x[0]), reverse=True)
    ocr_sidor = 0

    for sid_nr, page in enumerate(doc):
        text_pa_sida = page.get_text("text").strip()

        # Om sidan saknar textlager -> skannad PDF -> kör bildmaskering/OCR
        if len(text_pa_sida) < 20 and ocr_fallback:
            ocr_sidor += 1
            pix = page.get_pixmap(dpi=200)
            pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            bearbetad = bearbeta_pil_bild(
                pil_img,
                mapping,
                maskera_ritningshuvud,
                huvud_bredd_pct,
                huvud_hojd_pct,
                extra_flags,
            )
            buf = io.BytesIO()
            bearbetad.save(buf, format="PNG")
            ny_sida = nytt_doc.new_page(width=page.rect.width, height=page.rect.height)
            ny_sida.insert_image(ny_sida.rect, stream=buf.getvalue())
        else:
            # Textbaserad PDF -> äkta redaction
            for original, pseudonym in sorterade_par:
                for rect in page.search_for(original):
                    page.add_redact_annot(
                        rect,
                        text=pseudonym,
                        fill=(1, 1, 1),
                        text_color=(0, 0, 0),
                        fontsize=8,
                    )
            if any(extra_flags.values()):
                for monster, ers, aktiv in [
                    (REGEX_EPOST, "[E-POST]", extra_flags["maskera_epost"]),
                    (REGEX_TELEFON, "[TELEFON]", extra_flags["maskera_telefon"]),
                    (REGEX_KOORDINAT, "[KOORDINAT]", extra_flags["maskera_koord"]),
                ]:
                    if aktiv:
                        for match in set(monster.findall(text_pa_sida)):
                            for rect in page.search_for(match):
                                page.add_redact_annot(rect, text=ers, fill=(1, 1, 1), text_color=(0, 0, 0), fontsize=8)

            page.apply_redactions()
            nytt_doc.insert_pdf(doc, from_page=sid_nr, to_page=sid_nr)

    if ocr_sidor > 0:
        varningar.append(f"{ocr_sidor} skannad(e) sida/sidor bearbetades som bild.")

    nytt_doc.set_metadata({})
    nytt_doc.save(ut_fil, garbage=4, deflate=True)
    nytt_doc.close()
    doc.close()
    return varningar


# --- 5. LÄCKAGESKANNING ---
def skanna_efter_lackage(fil: Path, original_nycklar: list[str]) -> list[str]:
    if not original_nycklar:
        return []
    suffix = fil.suffix.lower()
    text = ""

    try:
        if suffix == ".docx":
            doc = Document(fil)
            delar = [p.text for p in doc.paragraphs]
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        delar.append(c.text)
            text = "\n".join(delar)
        elif suffix == ".pdf":
            with fitz.open(fil) as doc:
                text = "\n".join(p.get_text("text") for p in doc)
        elif suffix == ".xlsx":
            wb = openpyxl.load_workbook(fil, data_only=True)
            delar = list(wb.sheetnames)
            for s in wb.worksheets:
                for row in s.iter_rows(values_only=True):
                    for val in row:
                        if isinstance(val, str):
                            delar.append(val)
            text = "\n".join(delar)
        elif suffix in {".txt", ".md"}:
            text = fil.read_text(encoding="utf-8")
        else:
            return []
    except Exception:
        return []

    hittade = []
    text_lower = text.lower()
    for nyckel in original_nycklar:
        if nyckel.lower() in text_lower:
            hittade.append(nyckel)
    return hittade


# ==========================================
# TKINTER-GRÄNSSNITT
# ==========================================

class AnonymiserareApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Dokumentanonymiserare – Word / PDF / TIF / Excel")
        self.root.geometry("900x720")
        self.root.minsize(780, 600)

        self.nyckelfil = Path("nyckel.json")
        self.mapping: dict[str, str] = {}
        self.ladda_nyckel(skapa_om_saknas=True)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self._bygg_flik_filer()
        self._bygg_flik_text()
        self._bygg_flik_nyckel()

    # --- FLIK 1: BEARBETA FILER ---
    def _bygg_flik_filer(self) -> None:
        flik = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(flik, text="1. Bearbeta filer")

        self.in_mapp_var = tk.StringVar()
        self.ut_mapp_var = tk.StringVar()
        self.lage_var = tk.StringVar(value="anonymisera")
        self.undermappar_var = tk.BooleanVar(value=False)

        self.maskera_huvud_var = tk.BooleanVar(value=True)
        self.huvud_bredd_var = tk.DoubleVar(value=30.0)
        self.huvud_hojd_var = tk.DoubleVar(value=18.0)
        self.ocr_pdf_var = tk.BooleanVar(value=True)

        self.maskera_epost_var = tk.BooleanVar(value=True)
        self.maskera_tel_var = tk.BooleanVar(value=False)
        self.maskera_koord_var = tk.BooleanVar(value=False)
        self.lackageskanning_var = tk.BooleanVar(value=True)

        # Mappar
        mapp_ram = ttk.LabelFrame(flik, text="Mappar", padding=8)
        mapp_ram.pack(fill="x", pady=(0, 8))

        rad1 = ttk.Frame(mapp_ram)
        rad1.pack(fill="x", pady=2)
        ttk.Label(rad1, text="In-mapp:", width=10).pack(side="left")
        ttk.Entry(rad1, textvariable=self.in_mapp_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(rad1, text="Bläddra...", command=self._valj_in_mapp).pack(side="left")

        rad2 = ttk.Frame(mapp_ram)
        rad2.pack(fill="x", pady=2)
        ttk.Label(rad2, text="Ut-mapp:", width=10).pack(side="left")
        ttk.Entry(rad2, textvariable=self.ut_mapp_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(rad2, text="Bläddra...", command=self._valj_ut_mapp).pack(side="left", padx=(0, 5))
        ttk.Button(rad2, text="Öppna ut-mapp", command=self._oppna_ut_mapp).pack(side="left")

        # Inställningar
        inst_ram = ttk.LabelFrame(flik, text="Inställningar", padding=8)
        inst_ram.pack(fill="x", pady=(0, 8))

        rad_lage = ttk.Frame(inst_ram)
        rad_lage.pack(fill="x", pady=2)
        ttk.Radiobutton(rad_lage, text="Anonymisera (inför uppladdning)", variable=self.lage_var, value="anonymisera").pack(side="left", padx=(0, 15))
        ttk.Radiobutton(rad_lage, text="Återställ (av-anonymisera)", variable=self.lage_var, value="aterstall").pack(side="left", padx=(0, 15))
        ttk.Checkbutton(rad_lage, text="Inkludera undermappar", variable=self.undermappar_var).pack(side="left")

        rad_ritning = ttk.Frame(inst_ram)
        rad_ritning.pack(fill="x", pady=(6, 2))
        ttk.Checkbutton(
            rad_ritning,
            text="Vitmåla ritningshuvud (TIF & skannad PDF)",
            variable=self.maskera_huvud_var,
            command=self._uppdatera_huvud_falt,
        ).pack(side="left", padx=(0, 10))
        ttk.Label(rad_ritning, text="Bredd %:").pack(side="left")
        self.bredd_spin = ttk.Spinbox(rad_ritning, from_=5, to=100, increment=1, textvariable=self.huvud_bredd_var, width=5)
        self.bredd_spin.pack(side="left", padx=(2, 8))
        ttk.Label(rad_ritning, text="Höjd %:").pack(side="left")
        self.hojd_spin = ttk.Spinbox(rad_ritning, from_=5, to=50, increment=1, textvariable=self.huvud_hojd_var, width=5)
        self.hojd_spin.pack(side="left", padx=(2, 10))
        ttk.Button(rad_ritning, text="Förhandsgranska TIF...", command=self._forhandsgranska_tif).pack(side="left")

        rad_extra = ttk.Frame(inst_ram)
        rad_extra.pack(fill="x", pady=(6, 2))
        ttk.Checkbutton(rad_extra, text="OCR/maskering på skannade PDF:er", variable=self.ocr_pdf_var).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(rad_extra, text="Maskera e-post", variable=self.maskera_epost_var).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(rad_extra, text="Maskera telefon", variable=self.maskera_tel_var).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(rad_extra, text="Maskera koordinater", variable=self.maskera_koord_var).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(rad_extra, text="Läckageskanning", variable=self.lackageskanning_var).pack(side="left")

        # Kör-rad
        rad_kor = ttk.Frame(flik)
        rad_kor.pack(fill="x", pady=6)
        self.kor_btn = ttk.Button(rad_kor, text="▶ Kör bearbetning", command=self._starta_filbearbetning)
        self.kor_btn.pack(side="left")
        self.progress = ttk.Progressbar(rad_kor, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(10, 0))

        # Färgkodad logg
        ttk.Label(flik, text="Logg och kontrollrapport:").pack(anchor="w", pady=(4, 2))
        self.logg_text = scrolledtext.ScrolledText(flik, height=16, state="disabled", font=("Consolas", 9))
        self.logg_text.pack(fill="both", expand=True)
        self.logg_text.tag_configure("ok", foreground="#1b7a1b")
        self.logg_text.tag_configure("varning", foreground="#b36b00")
        self.logg_text.tag_configure("fel", foreground="#c00000")
        self.logg_text.tag_configure("info", foreground="#222222")

    def _uppdatera_huvud_falt(self) -> None:
        state = "normal" if self.maskera_huvud_var.get() else "disabled"
        self.bredd_spin.configure(state=state)
        self.hojd_spin.configure(state=state)

    def _valj_in_mapp(self) -> None:
        mapp = filedialog.askdirectory(title="Välj mapp med dokument")
        if mapp:
            self.in_mapp_var.set(mapp)
            if not self.ut_mapp_var.get():
                self.ut_mapp_var.set(str(Path(mapp) / "ut_anonymiserat"))

    def _valj_ut_mapp(self) -> None:
        mapp = filedialog.askdirectory(title="Välj ut-mapp")
        if mapp:
            self.ut_mapp_var.set(mapp)

    def _oppna_ut_mapp(self) -> None:
        ut_mapp = Path(self.ut_mapp_var.get().strip())
        if ut_mapp.is_dir():
            os.startfile(ut_mapp)
        else:
            messagebox.showinfo("Mapp saknas", "Ut-mappen finns inte ännu. Kör bearbetningen först.")

    def _forhandsgranska_tif(self) -> None:
        in_mapp = Path(self.in_mapp_var.get().strip())
        tif_fil = None
        if in_mapp.is_dir():
            tifs = list(in_mapp.glob("*.tif")) + list(in_mapp.glob("*.tiff"))
            if tifs:
                tif_fil = tifs[0]

        if not tif_fil:
            valt = filedialog.askopenfilename(title="Välj en TIF-ritning att förhandsgranska", filetypes=[("TIF-ritningar", "*.tif *.tiff")])
            if not valt:
                return
            tif_fil = Path(valt)

        try:
            bredd_pct = float(self.huvud_bredd_var.get())
            hojd_pct = float(self.huvud_hojd_var.get())
            img = Image.open(tif_fil).convert("RGB")
        except Exception as e:
            messagebox.showerror("Fel", f"Kunde inte öppna ritningen:\n{e}")
            return

        # Skala ned bilden så den passar i förhandsgranskningsfönstret
        max_w, max_h = 850, 600
        img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
        w, h = img.size

        draw = ImageDraw.Draw(img)
        x0 = int(w * (1.0 - bredd_pct / 100.0))
        y0 = int(h * (1.0 - hojd_pct / 100.0))
        draw.rectangle([x0, y0, w - 1, h - 1], outline="red", width=3)

        win = tk.Toplevel(self.root)
        win.title(f"Förhandsgranskning: {tif_fil.name} (Röd ram = vitmålas)")
        tk_img = ImageTk.PhotoImage(img)
        lbl = ttk.Label(win, image=tk_img)
        lbl.image = tk_img  # Behåll referens
        lbl.pack(padx=10, pady=10)
        ttk.Label(win, text=f"Maskerat område: {bredd_pct:.0f} % bredd × {hojd_pct:.0f} % höjd i nedre högra hörnet.").pack(pady=(0, 10))

    def _logga(self, meddelande: str, tag: str = "info") -> None:
        def _skriv() -> None:
            self.logg_text.configure(state="normal")
            self.logg_text.insert("end", meddelande + "\n", tag)
            self.logg_text.see("end")
            self.logg_text.configure(state="disabled")
        self.root.after(0, _skriv)

    def _starta_filbearbetning(self) -> None:
        in_mapp = Path(self.in_mapp_var.get().strip())
        ut_mapp = Path(self.ut_mapp_var.get().strip())

        if not in_mapp.is_dir():
            messagebox.showerror("Fel", "Välj en giltig in-mapp först.")
            return
        if not self.mapping:
            messagebox.showwarning("Tom nyckellista", "Lägg till minst ett namnpar under fliken '3. Nyckellista' först.")
            return

        self.kor_btn.configure(state="disabled")
        self.logg_text.configure(state="normal")
        self.logg_text.delete("1.0", "end")
        self.logg_text.configure(state="disabled")

        inst = {
            "aterstall": self.lage_var.get() == "aterstall",
            "rekursivt": self.undermappar_var.get(),
            "maskera_huvud": self.maskera_huvud_var.get(),
            "bredd_pct": float(self.huvud_bredd_var.get()),
            "hojd_pct": float(self.huvud_hojd_var.get()),
            "ocr_pdf": self.ocr_pdf_var.get(),
            "lackageskanning": self.lackageskanning_var.get(),
            "extra_flags": {
                "maskera_epost": self.maskera_epost_var.get() and self.lage_var.get() != "aterstall",
                "maskera_telefon": self.maskera_tel_var.get() and self.lage_var.get() != "aterstall",
                "maskera_koord": self.maskera_koord_var.get() and self.lage_var.get() != "aterstall",
            },
        }

        threading.Thread(
            target=self._kor_filer_i_trad,
            args=(in_mapp, ut_mapp, inst),
            daemon=True,
        ).start()

    def _kor_filer_i_trad(self, in_mapp: Path, ut_mapp: Path, inst: dict) -> None:
        filer = [
            f for f in (in_mapp.rglob("*") if inst["rekursivt"] else in_mapp.glob("*"))
            if f.is_file() and f.suffix.lower() in STODDA_FORMAT and ut_mapp not in f.parents
        ]

        if not filer:
            self._logga("Inga stödda filer (.docx, .pdf, .tif, .xlsx, .txt, .md) hittades.", "varning")
            self.root.after(0, lambda: self.kor_btn.configure(state="normal"))
            return

        self.root.after(0, lambda: self.progress.configure(maximum=len(filer), value=0))
        aterstall = inst["aterstall"]
        aktiv_mapping = {v: k for k, v in self.mapping.items()} if aterstall else self.mapping
        extra_flags = inst["extra_flags"]

        self._logga(f"{'Återställer' if aterstall else 'Anonymiserar'} {len(filer)} filer...\n", "info")

        ut_filer: list[Path] = []
        lyckade = 0

        for idx, fil in enumerate(filer, start=1):
            mal_mapp = ut_mapp / fil.parent.relative_to(in_mapp)
            mal_mapp.mkdir(parents=True, exist_ok=True)
            nytt_namn = ersatt_text(fil.name, aktiv_mapping)
            ut_fil = mal_mapp / nytt_namn
            suffix = fil.suffix.lower()

            try:
                varningar = []
                if suffix == ".docx":
                    varningar = bearbeta_docx(fil, ut_fil, aktiv_mapping, not aterstall, extra_flags)
                elif suffix == ".xlsx":
                    varningar = bearbeta_xlsx(fil, ut_fil, aktiv_mapping, not aterstall, extra_flags)
                elif suffix == ".pdf":
                    varningar = bearbeta_pdf(
                        fil, ut_fil, aktiv_mapping, inst["ocr_pdf"],
                        inst["maskera_huvud"] and not aterstall,
                        inst["bredd_pct"], inst["hojd_pct"], extra_flags,
                    )
                elif suffix in {".tif", ".tiff"}:
                    varningar = bearbeta_tif(
                        fil, ut_fil, aktiv_mapping,
                        inst["maskera_huvud"] and not aterstall,
                        inst["bredd_pct"], inst["hojd_pct"], extra_flags,
                    )
                elif suffix in {".txt", ".md"}:
                    text = fil.read_text(encoding="utf-8")
                    ut_fil.write_text(ersatt_text(text, aktiv_mapping, **extra_flags), encoding="utf-8")

                self._logga(f"[OK] {fil.name}  ->  {ut_fil.name}", "ok")
                for v in varningar:
                    self._logga(f"     [OBS] {v}", "varning")
                ut_filer.append(ut_fil)
                lyckade += 1
            except Exception as e:
                self._logga(f"[FEL] {fil.name}: {e}", "fel")

            self.root.after(0, lambda v=idx: self.progress.configure(value=v))

        # Läckageskanning efter anonymisering
        if inst["lackageskanning"] and not aterstall and ut_filer:
            self._logga("\n--- Kör läckageskanning på ut-filer ---", "info")
            original_nycklar = list(self.mapping.keys())
            lackage_hittat = False
            for uf in ut_filer:
                traffar = skanna_efter_lackage(uf, original_nycklar)
                if traffar:
                    lackage_hittat = True
                    self._logga(f"[LÄCKAGE] {uf.name} innehåller fortfarande: {', '.join(traffar)}", "fel")
            if not lackage_hittat:
                self._logga("[OK] Inga kvarvarande nyckelord hittades i textlagren.", "ok")

        self._logga(f"\nKlar! {lyckade}/{len(filer)} filer sparade i: {ut_mapp}", "info")
        self.root.after(0, lambda: self.kor_btn.configure(state="normal"))

    # --- FLIK 2: TEXT & AI-SVAR ---
    def _bygg_flik_text(self) -> None:
        flik = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(flik, text="2. Text & AI-svar")

        ttk.Label(flik, text="Klistra in text (t.ex. AI-svar eller textstycke):").pack(anchor="w")
        self.text_in = scrolledtext.ScrolledText(flik, height=11, font=("Segoe UI", 10))
        self.text_in.pack(fill="both", expand=True, pady=(2, 8))

        knappar = ttk.Frame(flik)
        knappar.pack(fill="x", pady=4)
        ttk.Button(knappar, text="Återställ text (Pseudonym -> Riktigt namn)", command=lambda: self._bearbeta_textruta(True)).pack(side="left", padx=(0, 8))
        ttk.Button(knappar, text="Anonymisera text (Riktigt namn -> Pseudonym)", command=lambda: self._bearbeta_textruta(False)).pack(side="left", padx=(0, 8))
        ttk.Button(knappar, text="Rensa", command=self._rensa_textrutor).pack(side="left")
        ttk.Button(knappar, text="Kopiera resultat", command=self._kopiera_resultat).pack(side="right")

        ttk.Label(flik, text="Resultat:").pack(anchor="w", pady=(8, 2))
        self.text_ut = scrolledtext.ScrolledText(flik, height=11, font=("Segoe UI", 10))
        self.text_ut.pack(fill="both", expand=True)

    def _bearbeta_textruta(self, aterstall: bool) -> None:
        in_text = self.text_in.get("1.0", "end-1c")
        aktiv = {v: k for k, v in self.mapping.items()} if aterstall else self.mapping
        resultat = ersatt_text(in_text, aktiv)
        self.text_ut.delete("1.0", "end")
        self.text_ut.insert("1.0", resultat)

    def _rensa_textrutor(self) -> None:
        self.text_in.delete("1.0", "end")
        self.text_ut.delete("1.0", "end")

    def _kopiera_resultat(self) -> None:
        text = self.text_ut.get("1.0", "end-1c")
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("Kopierat", "Resultatet har kopierats till urklipp.")

    # --- FLIK 3: NYCKELLISTA ---
    def _bygg_flik_nyckel(self) -> None:
        flik = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(flik, text="3. Nyckellista (nyckel.json)")

        top = ttk.Frame(flik)
        top.pack(fill="x", pady=(0, 8))
        self.nyckel_sokvag_lbl = ttk.Label(top, text=f"Aktiv fil: {self.nyckelfil.resolve()}")
        self.nyckel_sokvag_lbl.pack(side="left")
        ttk.Button(top, text="Öppna annan nyckelfil...", command=self._valj_nyckelfil).pack(side="right")

        self.tree = ttk.Treeview(flik, columns=("original", "pseudonym"), show="headings", height=14)
        self.tree.heading("original", text="Riktigt namn (Original)")
        self.tree.heading("pseudonym", text="Ersättningsnamn (Pseudonym)")
        self.tree.column("original", width=380)
        self.tree.column("pseudonym", width=380)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._fyll_fran_tabell)
        self._uppdatera_nyckeltabell()

        ttk.Label(flik, text="Tips: Dubbelklicka på en rad för att redigera den. Versaler/gemener anpassas automatiskt.").pack(anchor="w", pady=(4, 2))

        inmatning = ttk.Frame(flik)
        inmatning.pack(fill="x", pady=6)

        ttk.Label(inmatning, text="Original:").pack(side="left")
        self.orig_entry = ttk.Entry(inmatning, width=28)
        self.orig_entry.pack(side="left", padx=(4, 10))

        ttk.Label(inmatning, text="Pseudonym:").pack(side="left")
        self.pseudo_entry = ttk.Entry(inmatning, width=28)
        self.pseudo_entry.pack(side="left", padx=(4, 10))

        ttk.Button(inmatning, text="Lägg till / Uppdatera", command=self._lagg_till_nyckel).pack(side="left", padx=4)
        ttk.Button(inmatning, text="Ta bort markerad", command=self._ta_bort_nyckel).pack(side="right")

    def _fyll_fran_tabell(self, _event) -> None:
        markerade = self.tree.selection()
        if not markerade:
            return
        orig, pseudo = self.tree.item(markerade[0], "values")
        self.orig_entry.delete(0, "end")
        self.orig_entry.insert(0, orig)
        self.pseudo_entry.delete(0, "end")
        self.pseudo_entry.insert(0, pseudo)

    def ladda_nyckel(self, skapa_om_saknas: bool = False) -> None:
        if not self.nyckelfil.exists():
            if skapa_om_saknas:
                self.mapping = {"Exempel AB": "Företag A"}
                self.spara_nyckel()
            else:
                self.mapping = {}
            return
        try:
            self.mapping = json.loads(self.nyckelfil.read_text(encoding="utf-8"))
        except Exception as e:
            messagebox.showerror("Fel vid läsning", f"Kunde inte läsa {self.nyckelfil}:\n{e}")
            self.mapping = {}

    def spara_nyckel(self) -> None:
        self.nyckelfil.write_text(json.dumps(self.mapping, ensure_ascii=False, indent=2), encoding="utf-8")

    def _uppdatera_nyckeltabell(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for orig, pseudo in sorted(self.mapping.items()):
            self.tree.insert("", "end", values=(orig, pseudo))

    def _lagg_till_nyckel(self) -> None:
        orig = self.orig_entry.get().strip()
        pseudo = self.pseudo_entry.get().strip()
        if not orig or not pseudo:
            messagebox.showwarning("Saknas text", "Fyll i både original och pseudonym.")
            return
        self.mapping[orig] = pseudo
        self.spara_nyckel()
        self._uppdatera_nyckeltabell()
        self.orig_entry.delete(0, "end")
        self.pseudo_entry.delete(0, "end")

    def _ta_bort_nyckel(self) -> None:
        markerade = self.tree.selection()
        if not markerade:
            return
        for item in markerade:
            orig = self.tree.item(item, "values")[0]
            self.mapping.pop(orig, None)
        self.spara_nyckel()
        self._uppdatera_nyckeltabell()

    def _valj_nyckelfil(self) -> None:
        fil = filedialog.askopenfilename(title="Välj nyckelfil", filetypes=[("JSON-filer", "*.json")])
        if fil:
            self.nyckelfil = Path(fil)
            self.ladda_nyckel()
            self._uppdatera_nyckeltabell()
            self.nyckel_sokvag_lbl.configure(text=f"Aktiv fil: {self.nyckelfil.resolve()}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AnonymiserareApp(root)
    root.mainloop()
