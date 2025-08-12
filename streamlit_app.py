# app.py
import io
import os
import re
import zipfile
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors


# ============================================================
# Konfigūracija
# ============================================================
st.set_page_config(page_title="Mokinių PDF generatorius", layout="centered")
st.title("🎓 PDF generatorius kiekvienam mokiniui iš Excel")

st.write(
    "Įkelkite suvestinę tokio formato, kaip pavyzdyje: viršuje informacinės eilutės, "
    "antraštė su „Eil. Nr.“ ir „Pavardė, vardas“, žemiau — dalykų pavadinimai."
)

# Pastovi šrifto vieta (1) – kaip nurodėte
USER_FIXED_FONT_PATH = "DejaVuSerif.ttf"
# Pastovi šrifto vieta (2) – bandymas rasti šalia app.py
LOCAL_FONT_PATH = os.path.join(os.path.dirname(__file__), "DejaVuSerif.ttf")

# --- Filtrai pasiekimų lygiui / stulpeliams ---
ACH_LVL_COL_RE = re.compile(r"(?i)pasiek\w*\s*lyg\w*")  # stulpelių pavadinimams
ACH_LVL_VAL_RE = re.compile(r"(?i)^(slenkstinis|patenkinamas|pagrindinis|aukštesnysis)\b")

# --- Valymo šablonai „PR“ / „IN“ žymoms pašalinti ---
TAG_PR_IN_RE = re.compile(r"(?i)\b(?:PR|IN)\b")
EXTRA_SEP_RE = re.compile(r"\s*[(){}\[\]/|,;.-]\s*")  # tvarkyti likusius skyriklius po tagų trynimo


def register_fixed_font() -> str:
    """
    Registruoja DejaVuSerif.ttf iš fiksuoto kelio arba iš programos aplanko.
    Jei nepavyksta — grįžta į Helvetica.
    """
    font_candidates = [USER_FIXED_FONT_PATH, LOCAL_FONT_PATH]
    for p in font_candidates:
        if p and os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("DejaVuSerifFixed", p))
                return "DejaVuSerifFixed"
            except Exception as e:
                st.warning(f"Šrifto registruoti nepavyko iš '{p}': {e}")
    st.warning("Nerastas DejaVuSerif.ttf — bus naudojama Helvetica (lietuviškos raidės gali būti neteisingos).")
    return "Helvetica"


FONT_NAME = register_fixed_font()

# Įkėlimas
excel_file = st.file_uploader("📄 Excel (.xlsx, .xls)", type=["xlsx", "xls"])

# Viršeliui/atributams
col_a, col_b, col_c = st.columns(3)
with col_a:
    akadem_metai_input = st.text_input("Mokslo metai (pvz., 2024–2025)", value="")
with col_b:
    override_school = st.text_input("Mokykla (jei reikia)", value="")
with col_c:
    override_class = st.text_input("Klasė (jei reikia)", value="")


# ============================================================
# Pagalbinės funkcijos
# ============================================================
def try_extract_school_and_class(df_raw: pd.DataFrame):
    mokykla, klase = "", ""
    nrows = min(12, df_raw.shape[0])
    ncols = min(12, df_raw.shape[1])
    for r in range(nrows):
        for c in range(ncols):
            val = df_raw.iat[r, c]
            if isinstance(val, str):
                txt = val.strip()
                low = txt.lower()
                if "mokykla" in low and ":" in txt:
                    mokykla = txt.split(":", 1)[1].strip()
                if ("klasė" in low or "klase" in low) and ":" in txt:
                    klase = txt.split(":", 1)[1].strip()
    return mokykla, klase


def try_extract_academic_year(df_raw: pd.DataFrame) -> str:
    pat_years = re.compile(r"(20\d{2})\s*[–-]\s*(20\d{2})")
    nrows = min(12, df_raw.shape[0])
    ncols = min(12, df_raw.shape[1])
    candidate = ""
    for r in range(nrows):
        for c in range(ncols):
            val = df_raw.iat[r, c]
            if isinstance(val, str):
                txt = val.strip()
                m = pat_years.search(txt)
                if m:
                    return f"{m.group(1)}–{m.group(2)}"
                if "mokslo metai" in txt.lower():
                    m2 = pat_years.search(txt)
                    if m2:
                        return f"{m2.group(1)}–{m2.group(2)}"
                    candidate = txt
    return candidate


def find_header_rows(df0: pd.DataFrame):
    max_scan = min(40, len(df0))
    for i in range(max_scan):
        c0 = df0.iat[i, 0]
        c1 = df0.iat[i, 1] if df0.shape[1] > 1 else None
        c0s = (str(c0).strip() if pd.notna(c0) else "").lower()
        c1s = (str(c1).strip() if pd.notna(c1) else "").lower()
        if c0s.startswith("eil") and ("pavard" in c1s or "vard" in c1s):
            return i, i + 1
    return None, None


def parse_excel_to_df(excel_bytes: bytes):
    """
    Skaito Excel kaip df0 be antraščių, iškerpa dalį su mokiniais.
    Grąžina: df_students (['Eil. Nr.', 'Pavardė, vardas', <dalykai...>]), meta(dict).
    """
    df0 = pd.read_excel(io.BytesIO(excel_bytes), header=None)
    mokykla, klase = try_extract_school_and_class(df0)
    akadem_metai_auto = try_extract_academic_year(df0)

    header_row, subjects_row = find_header_rows(df0)
    if header_row is None:
        raise ValueError("Nerasta antraštės eilutė: turi būti 'Eil. Nr.' ir 'Pavardė, vardas'.")

    # Stulpelių sudarymas
    columns = ["Eil. Nr.", "Pavardė, vardas"]
    for j in range(2, df0.shape[1]):
        name = df0.iat[subjects_row, j]
        name = "" if pd.isna(name) else str(name).strip()
        if not name:
            name = f"Dalykas_{j-1}"
        columns.append(name)

    # Pirmoji duomenų eilutė — po dalykų pavadinimų; ieškome, kur prasideda numeracija
    data_start = subjects_row + 1
    while data_start < len(df0):
        v = df0.iat[data_start, 0]
        if pd.notna(v):
            try:
                int(str(v).strip())
                break
            except Exception:
                pass
        data_start += 1

    data = df0.iloc[data_start:, : len(columns)].copy()
    data.columns = columns
    data = data.dropna(how="all")
    data = data[data["Pavardė, vardas"].notna()]

    # Paliekame tik tuos dalykus, kur yra bent viena reikšmė
    keep_cols = ["Eil. Nr.", "Pavardė, vardas"]
    for col in data.columns:
        if col in keep_cols:
            continue
        if data[col].notna().any():
            keep_cols.append(col)
    data = data[keep_cols]

    meta = {"mokykla": mokykla, "klase": klase, "akadem_metai": akadem_metai_auto}
    return data.reset_index(drop=True), meta


# --- klasės logika: ar tai baigiamoji (IV ar 12) ---
FINAL_CLASS_RE = re.compile(r"(?i)^\s*IV\s*[A-ZĄČĘĖĮŠŲŪŽ]?\s*$")
HAS_12_RE = re.compile(r"(?i)\b12\b")


def is_final_class(klase: str) -> bool:
    if not isinstance(klase, str):
        return False
    k = klase.strip()
    if FINAL_CLASS_RE.match(k):
        return True
    if HAS_12_RE.search(k):
        return True
    return False


def strip_accents_lower(s: str) -> str:
    """Normalizuoja: nuima diakritikus ir mažosios raidės palyginimui."""
    if not isinstance(s, str):
        return ""
    n = unicodedata.normalize("NFD", s)
    n = "".join(ch for ch in n if unicodedata.category(ch) != "Mn")
    return n.lower().strip()


def clean_value_remove_tags(val: str) -> str:
    """
    Pašalina žymas PR/IN (atskirai stovinčias) ir sutvarko likusius skyriklius.
    Pvz.: '8 PR' -> '8'; 'įsk (IN)' -> 'įsk'.
    """
    s = str(val).strip()
    if not s:
        return s
    # pašalinam atskirus PR/IN žodžius
    s = TAG_PR_IN_RE.sub("", s)
    # sutvarkom likusius skyriklius tarp tuščių dalių
    # pvz. '8  / ' -> '8'; 'įsk ()' -> 'įsk'
    s = re.sub(r"\s{2,}", " ", s)
    s = s.strip(" ()[]{}-/|,;.")
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s


def make_student_pdf(
    buf: io.BytesIO,
    font_name: str,
    student_name: str,
    klasė: str,
    akademiniai_metai: str,
    school: str,
    subjects_dict: dict
):
    """
    Kuria vieno mokinio PDF į nurodytą buferį.
    Išmeta eilutes su tuščiais įrašais ir su pasiekimų lygio reikšmėmis.
    Po lentelės prideda sakinį apie perkėlimą arba baigimą (pagal klasę).
    Taip pat pašalina 'PR'/'IN' žymes iš reikšmių.
    """
    styles = getSampleStyleSheet()
    for k in list(styles.byName.keys()):
        styles.byName[k].fontName = font_name

    title_style = ParagraphStyle(
        "TitleLT",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=16,
        leading=20,
        spaceAfter=10,
        alignment=1,
    )
    normal = ParagraphStyle(
        "NormalLT",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=10,
        leading=14,
    )

    doc = SimpleDocTemplate(
        buf, pagesize=A4, rightMargin=2 * cm, leftMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm
    )

    story = []
    story.append(Paragraph("Mokinio pasiekimų įrašas", title_style))
    if school:
        story.append(Paragraph(f"Mokykla: <b>{school}</b>", normal))
    if klasė:
        story.append(Paragraph(f"Klasė: <b>{klasė}</b>", normal))
    if akademiniai_metai:
        story.append(Paragraph(f"Mokslo metai: <b>{akademiniai_metai}</b>", normal))
    story.append(Paragraph(f"Mokinys: <b>{student_name}</b>", normal))
    story.append(Spacer(1, 10))

    # Lentelė: Dalykas | Įvertinimas
    data_tbl = [["Dalykas", "Įvertinimas"]]
    for subj, val in subjects_dict.items():
        if pd.isna(val):
            continue
        sval = str(val).strip()
        if not sval or sval.lower() == "nan":
            continue
        if ACH_LVL_VAL_RE.search(sval):
            continue
        sval = clean_value_remove_tags(sval)
        if not sval:
            continue
        data_tbl.append([subj, sval])

    tbl = Table(data_tbl, colWidths=[10 * cm, 4 * cm])
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (1, 1), (1, -1), "CENTER"),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ]
        )
    )
    story.append(tbl)

    # --- privalomas sakinys po lentelės ---
    story.append(Spacer(1, 10))
    if is_final_class(klasė or ""):
        story.append(Paragraph("Baigė vidurinio ugdymo programą.", normal))
    else:
        story.append(Paragraph("Direktoriaus įsakymu perkeltas/a į aukštesnę klasę.", normal))

    doc.build(story)


# ============================================================
# Veiksmas
# ============================================================
if excel_file is not None:
    try:
        excel_bytes = excel_file.read()
        df_students, meta = parse_excel_to_df(excel_bytes)

        # Pašaliname stulpelius su „pasiekimų lygis“ pavadinimais
        df_students = df_students[[c for c in df_students.columns
                                   if not ACH_LVL_COL_RE.search(str(c) or "")]]

        # Pašaliname eilutę „Klasės pažangumas“, jei ji kaip "mokinio vardas"
        # (lyginame be diakritinių ženklų, neatsižvelgiant į raidžių dydį)
        name_norm = df_students["Pavardė, vardas"].astype(str).map(strip_accents_lower)
        df_students = df_students[~name_norm.eq(strip_accents_lower("Klasės pažangumas"))].reset_index(drop=True)

        st.subheader("Trumpa peržiūra")
        st.dataframe(df_students.head(10), use_container_width=True)

        school = override_school or meta.get("mokykla", "")
        klase = override_class or meta.get("klase", "")
        akadem_metai = akadem_metai_input or meta.get("akadem_metai", "")

        if st.button("🚀 Generuoti PDF kiekvienam mokiniui"):
            if df_students.empty:
                st.error("Nerasta duomenų eilučių su mokinių vardais.")
            else:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for _, row in df_students.iterrows():
                        student_name = str(row.get("Pavardė, vardas", "")).strip()
                        if not student_name or student_name.lower() == "nan":
                            continue
                        # apsauga dar kartą (jei pavadinimas netikėtai kitas atėjimo taške)
                        if strip_accents_lower(student_name) == strip_accents_lower("Klasės pažangumas"):
                            continue

                        # Filtruoti dalykai / reikšmės: ne-NaN, ne tuščia, ne pasiekimų lygiai; taip pat valome PR/IN
                        subjects = {}
                        for col in df_students.columns:
                            if col in ("Eil. Nr.", "Pavardė, vardas"):
                                continue
                            v = row[col]
                            if pd.isna(v):
                                continue
                            sv = str(v).strip()
                            if not sv or sv.lower() == "nan":
                                continue
                            if ACH_LVL_VAL_RE.search(sv):
                                continue
                            sv = clean_value_remove_tags(sv)
                            if not sv:
                                continue
                            subjects[col] = sv

                        pdf_bytes = io.BytesIO()
                        make_student_pdf(
                            buf=pdf_bytes,
                            font_name=FONT_NAME,
                            student_name=student_name,
                            klasė=klase,
                            akademiniai_metai=akadem_metai,
                            school=school,
                            subjects_dict=subjects
                        )
                        pdf_bytes.seek(0)
                        safe_name = (
                            student_name.replace("/", "-")
                            .replace("\\", "-")
                            .replace(":", "-")
                            .replace("*", "-")
                            .replace("?", "")
                            .replace('"', "")
                            .replace("<", "(")
                            .replace(">", ")")
                            .replace("|", "-")
                        )
                        zf.writestr(f"{safe_name}.pdf", pdf_bytes.read())

                zip_buf.seek(0)
                st.success("Parengta! Atsisiųskite ZIP su visais PDF.")
                st.download_button(
                    label="⬇️ Atsisiųsti visus PDF (ZIP)",
                    data=zip_buf.getvalue(),
                    file_name="mokiniu_pdf.zip",
                    mime="application/zip",
                )

    except Exception as e:
        st.error(f"Nepavyko apdoroti failo: {e}")
        st.exception(e)
else:
    st.info("Įkelkite Excel failą, kad pradėti.")
