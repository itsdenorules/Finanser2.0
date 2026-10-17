# -*- coding: utf-8 -*-
# ============================================================================
#  FİNANSER — Yeni Nesil On Muhasebe & Finansal Takip SaaS
# ============================================================================
#  Gereksinimler (requirements.txt):
#    streamlit>=1.30.0
#    google-generativeai
#    openpyxl
#    Pillow
# ============================================================================

import streamlit as st
import sqlite3
import hashlib
import uuid
import json
import os
import io
import re
from datetime import datetime, date
import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------------------------------------------------------------------
#  Sayfa Yapilandirmasi
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Finanser",
    page_icon="F",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
#  Sabitler
# ---------------------------------------------------------------------------
PAYTR_AYLIK_LINK = ""   # TODO: BURAYA PAYTR AYLIK ODEME LINKI EKLENECEK
PAYTR_YILLIK_LINK = ""  # TODO: BURAYA PAYTR YILLIK ODEME LINKI EKLENECEK

VALID_REF_CODES = ["finans101", "finans102", "finans103"]
ADMIN_EMAIL = "denizzdinc@icloud.com"
DB_PATH = "finanser.db"

GELIR_KATEGORILERI = [
    "Satis",
    "Hizmet",
    "Faiz Geliri",
    "Kira Geliri",
    "Komisyon",
    "Diger Gelir",
]
GIDER_KATEGORILERI = [
    "Kira",
    "Fatura / Aidat",
    "Personel",
    "Malzeme / Hammadde",
    "Ulasim",
    "Yemek",
    "Ofis Gideri",
    "Pazarlama / Reklam",
    "Vergi / Harc",
    "Sigorta",
    "Bakim / Onarim",
    "Iletisim",
    "Diger Gider",
]

# ---------------------------------------------------------------------------
#  Gemini API
# ---------------------------------------------------------------------------
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel("gemini-1.5-flash")


# ===================================================================
#  CSS — Tum gorunumu kontrol eden tek blok
# ===================================================================
def inject_css():
    css = """<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*{font-family:'Inter',sans-serif!important}
.stApp{background-color:#0D0D0D!important}
[data-testid="stSidebar"]{background-color:#111111!important;border-right:1px solid #1E1E1E!important}
[data-testid="stSidebar"] *{color:#F5F5F7!important}
header[data-testid="stHeader"]{background-color:#0D0D0D!important}
.stApp p,.stApp span,.stApp label,.stApp div,.stApp h1,.stApp h2,.stApp h3,.stApp h4{color:#F5F5F7}
.stButton>button,.stFormSubmitButton>button{
background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%)!important;
color:#0D0D0D!important;font-weight:700!important;border:none!important;
border-radius:16px!important;padding:10px 28px!important;
transition:all .2s ease!important;letter-spacing:.3px!important}
.stButton>button:hover,.stFormSubmitButton>button:hover{
opacity:.88!important;transform:translateY(-1px)!important}
.stButton>button:active,.stFormSubmitButton>button:active{
transform:translateY(0)!important;opacity:1!important}
.stDownloadButton>button{
background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%)!important;
color:#0D0D0D!important;font-weight:700!important;border:none!important;
border-radius:16px!important;padding:10px 28px!important;transition:all .2s ease!important}
.stDownloadButton>button:hover{opacity:.88!important;transform:translateY(-1px)!important}
.stTextInput>div>div>input,
.stNumberInput>div>div>input,
.stDateInput>div>div>input,
.stTextArea textarea{
background-color:#1A1A1A!important;color:#F5F5F7!important;
border:1px solid #2A2A2A!important;border-radius:12px!important}
.stTextInput>div>div>input:focus,
.stNumberInput>div>div>input:focus,
.stTextArea textarea:focus{
border-color:#D4D8C8!important;box-shadow:0 0 0 1px #D4D8C8!important}
.stSelectbox>div>div{
background-color:#1A1A1A!important;color:#F5F5F7!important;
border:1px solid #2A2A2A!important;border-radius:12px!important}
[data-baseweb="select"] *{color:#F5F5F7!important}
[data-baseweb="popover"]{background-color:#1A1A1A!important;border:1px solid #2A2A2A!important}
[data-baseweb="menu"]{background-color:#1A1A1A!important}
[role="option"]{color:#F5F5F7!important}
[role="option"]:hover{background-color:#2A2A2A!important}
[aria-selected="true"]{background-color:#333!important}
.InputInstructions{display:none!important}
[data-testid="InputInstructions"]{display:none!important}
footer{visibility:hidden!important}
#MainMenu{visibility:hidden!important}
[data-testid="stFileUploader"]{background-color:#161616!important;border-radius:12px!important;border:1px dashed #2A2A2A!important;padding:16px!important}
[data-testid="stFileUploader"] *{color:#F5F5F7!important}
[data-testid="stFileUploader"] button{
background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%)!important;
color:#0D0D0D!important;font-weight:600!important;border-radius:12px!important;border:none!important}
.stRadio>div{flex-direction:column;gap:2px}
.stRadio label{
background-color:transparent!important;padding:10px 14px!important;
border-radius:10px!important;cursor:pointer!important;transition:background .15s!important;
font-weight:500!important;font-size:14px!important}
.stRadio label:hover{background-color:#1E1E1E!important}
.stRadio label[data-checked="true"]{
background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%)!important;
color:#0D0D0D!important}
.stTabs [data-baseweb="tab-list"]{gap:8px;background:transparent!important;border-bottom:1px solid #1E1E1E!important}
.stTabs [data-baseweb="tab"]{
background-color:#161616!important;color:#999!important;
border-radius:10px 10px 0 0!important;padding:10px 22px!important;
font-weight:500!important;border:none!important}
.stTabs [aria-selected="true"]{
background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%)!important;
color:#0D0D0D!important;font-weight:700!important}
.stTabs [data-baseweb="tab-panel"]{background:transparent!important;padding-top:20px!important}
div[data-testid="stExpander"]{background-color:#161616!important;border:1px solid #1E1E1E!important;border-radius:12px!important}
div[data-testid="stExpander"] *{color:#F5F5F7!important}
.stAlert{border-radius:12px!important}
hr{border-color:#1E1E1E!important}
[data-testid="stForm"]{background-color:#161616!important;border:1px solid #1E1E1E!important;border-radius:16px!important;padding:24px!important}
</style>"""
    st.markdown(css, unsafe_allow_html=True)


# ===================================================================
#  Veritabani
# ===================================================================
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


@st.cache_resource
def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute(
        "CREATE TABLE IF NOT EXISTS users ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "ad TEXT NOT NULL,"
        "email TEXT UNIQUE NOT NULL,"
        "sifre_hash TEXT NOT NULL,"
        "rol TEXT DEFAULT 'user',"
        "paket TEXT DEFAULT 'Ucretsiz',"
        "kota INTEGER DEFAULT 5,"
        "kullanilan_belge INTEGER DEFAULT 0,"
        "referans_kodu TEXT,"
        "session_token TEXT,"
        "created_at TEXT DEFAULT (datetime('now','localtime')))"
    )
    c.execute(
        "CREATE TABLE IF NOT EXISTS transactions ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "user_id INTEGER NOT NULL,"
        "tarih TEXT,"
        "islem_tipi TEXT,"
        "kategori TEXT,"
        "firma TEXT,"
        "tutar REAL DEFAULT 0,"
        "kdv REAL DEFAULT 0,"
        "aciklama TEXT,"
        "belge_turu TEXT,"
        "kaynak TEXT DEFAULT 'Manuel',"
        "created_at TEXT DEFAULT (datetime('now','localtime')),"
        "FOREIGN KEY (user_id) REFERENCES users(id))"
    )
    conn.commit()
    conn.close()
    return True


# ===================================================================
#  Sifre & Token Yardimcilari
# ===================================================================
def hash_password(pw, salt=None):
    if salt is None:
        salt = os.urandom(16).hex()
    h = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt.encode("utf-8"), 100_000).hex()
    return f"{salt}${h}"


def verify_password(pw, stored):
    parts = stored.split("$")
    if len(parts) != 2:
        return False
    return hash_password(pw, parts[0]) == stored


def generate_token():
    return uuid.uuid4().hex


# ===================================================================
#  Kimlik Dogrulama
# ===================================================================
def register_user(ad, email, pw, ref_code=""):
    conn = get_db()
    c = conn.cursor()
    try:
        rol = "admin" if email.lower() == ADMIN_EMAIL else "user"
        token = generate_token()
        c.execute(
            "INSERT INTO users (ad,email,sifre_hash,rol,referans_kodu,session_token) VALUES (?,?,?,?,?,?)",
            (ad, email.lower(), hash_password(pw), rol, ref_code, token),
        )
        conn.commit()
        uid = c.lastrowid
        conn.close()
        return {"ok": True, "token": token, "id": uid}
    except sqlite3.IntegrityError:
        conn.close()
        return {"ok": False, "err": "Bu e-posta adresi zaten kayitli."}


def login_user(email, pw):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE email=?", (email.lower(),))
    row = c.fetchone()
    if row is None:
        conn.close()
        return {"ok": False, "err": "E-posta adresi bulunamadi."}
    if not verify_password(pw, row["sifre_hash"]):
        conn.close()
        return {"ok": False, "err": "Sifre hatali."}
    if email.lower() == ADMIN_EMAIL and row["rol"] != "admin":
        c.execute("UPDATE users SET rol='admin' WHERE id=?", (row["id"],))
    token = generate_token()
    c.execute("UPDATE users SET session_token=? WHERE id=?", (token, row["id"]))
    conn.commit()
    conn.close()
    return {"ok": True, "token": token, "id": row["id"]}


def get_user_by_token(token):
    if not token:
        return None
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE session_token=?", (token,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_info(uid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE id=?", (uid,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def logout_user(uid):
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE users SET session_token=NULL WHERE id=?", (uid,))
    conn.commit()
    conn.close()


# ===================================================================
#  Islem (Transaction) Fonksiyonlari
# ===================================================================
def add_transaction(uid, tarih, islem_tipi, kategori, firma, tutar, kdv, aciklama, belge_turu, kaynak):
    conn = get_db()
    c = conn.cursor()
    c.execute(
        "INSERT INTO transactions (user_id,tarih,islem_tipi,kategori,firma,tutar,kdv,aciklama,belge_turu,kaynak) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (uid, tarih, islem_tipi, kategori, firma, tutar, kdv, aciklama, belge_turu, kaynak),
    )
    conn.commit()
    conn.close()


def increment_belge(uid):
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE users SET kullanilan_belge=kullanilan_belge+1 WHERE id=?", (uid,))
    conn.commit()
    conn.close()


def get_user_transactions(uid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM transactions WHERE user_id=? ORDER BY created_at DESC", (uid,))
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_user_stats(uid):
    conn = get_db()
    c = conn.cursor()
    c.execute(
        "SELECT "
        "COALESCE(SUM(CASE WHEN islem_tipi='Gelir' THEN tutar ELSE 0 END),0) as toplam_gelir,"
        "COALESCE(SUM(CASE WHEN islem_tipi='Gider' THEN tutar ELSE 0 END),0) as toplam_gider,"
        "COUNT(*) as islem_sayisi "
        "FROM transactions WHERE user_id=?",
        (uid,),
    )
    row = c.fetchone()
    conn.close()
    return dict(row)


# ===================================================================
#  Admin Fonksiyonlari
# ===================================================================
def get_all_users():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users ORDER BY created_at DESC")
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_all_transactions_admin():
    conn = get_db()
    c = conn.cursor()
    c.execute(
        "SELECT t.*,u.ad,u.email FROM transactions t "
        "JOIN users u ON t.user_id=u.id ORDER BY t.created_at DESC LIMIT 200"
    )
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_user_package(uid, paket):
    kota = 5 if paket == "Ucretsiz" else 1500
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE users SET paket=?,kota=? WHERE id=?", (paket, kota, uid))
    conn.commit()
    conn.close()


def get_total_belge():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COALESCE(SUM(kullanilan_belge),0) as t FROM users")
    row = c.fetchone()
    conn.close()
    return row["t"]


# ===================================================================
#  Gemini ile Belge Isleme
# ===================================================================
def process_document_with_ai(file_bytes, mime_type):
    prompt = (
        "Bu bir mali belgedir (fatura, fis, banka dekontu veya banka ekstresi olabilir). "
        "Belgeyi dikkatlice analiz et ve asagidaki bilgileri JSON formatinda cikar:\n\n"
        '{"belge_turu":"Fatura / Fis / Banka Dekontu / Banka Ekstresi",'
        '"tarih":"GG.AA.YYYY",'
        '"firma":"Firma veya kisi adi",'
        '"kategori":"Uygun kategori",'
        '"tutar":0.00,'
        '"kdv":0.00,'
        '"islem_tipi":"Gelir veya Gider",'
        '"aciklama":"Kisa aciklama"}\n\n'
        "SADECE JSON dondur, baska hicbir sey yazma. "
        "Eger bir bilgiyi cikaramiyorsan mantikli bir tahmin yap."
    )
    try:
        resp = gemini_model.generate_content(
            [prompt, {"mime_type": mime_type, "data": file_bytes}]
        )
        text = resp.text.strip()
        if text.startswith("```"):
            text = re.sub(r"^```(?:json)?\s*", "", text)
            text = re.sub(r"\s*```$", "", text)
        data = json.loads(text)
        return {"ok": True, "data": data}
    except Exception as e:
        return {"ok": False, "err": str(e)}


# ===================================================================
#  Excel Uretimi
# ===================================================================
def generate_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Islemler"
    headers = [
        "Tarih", "Islem Tipi", "Kategori", "Firma / Kisi",
        "Tutar (TL)", "KDV (TL)", "Aciklama", "Belge Turu", "Kaynak",
    ]
    hfont = Font(bold=True, color="FFFFFF", size=11)
    hfill = PatternFill(start_color="0D0D0D", end_color="0D0D0D", fill_type="solid")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = hfont
        cell.fill = hfill
        cell.alignment = Alignment(horizontal="center")
    for ri, t in enumerate(rows, 2):
        ws.cell(row=ri, column=1, value=t.get("tarih", ""))
        ws.cell(row=ri, column=2, value=t.get("islem_tipi", ""))
        ws.cell(row=ri, column=3, value=t.get("kategori", ""))
        ws.cell(row=ri, column=4, value=t.get("firma", ""))
        ws.cell(row=ri, column=5, value=t.get("tutar", 0))
        ws.cell(row=ri, column=6, value=t.get("kdv", 0))
        ws.cell(row=ri, column=7, value=t.get("aciklama", ""))
        ws.cell(row=ri, column=8, value=t.get("belge_turu", ""))
        ws.cell(row=ri, column=9, value=t.get("kaynak", ""))
    for col in ws.columns:
        mx = 0
        letter = col[0].column_letter
        for cell in col:
            if cell.value:
                mx = max(mx, len(str(cell.value)))
        ws.column_dimensions[letter].width = mx + 5
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ===================================================================
#  Yardimci Fonksiyonlar
# ===================================================================
def get_greeting(name):
    h = datetime.now().hour
    if 6 <= h < 12:
        return f"Gunaydin, {name}"
    elif 12 <= h < 18:
        return f"Merhaba, {name}"
    return f"Iyi aksamlar, {name}"


def format_tl(val):
    try:
        val = float(val)
    except (ValueError, TypeError):
        return "0,00 TL"
    s = f"{val:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " TL"


def format_date_tr(dt_str):
    if not dt_str:
        return "-"
    try:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y"):
            try:
                d = datetime.strptime(dt_str, fmt)
                return d.strftime("%d.%m.%Y")
            except ValueError:
                continue
        return dt_str
    except Exception:
        return dt_str


# ===================================================================
#  HTML Karti Olusturucular
# ===================================================================
def card_dark(content_html):
    return (
        '<div style="background:#161616;border:1px solid #1E1E1E;border-radius:16px;'
        'padding:24px;margin-bottom:16px;">'
        + content_html
        + "</div>"
    )


def card_gradient(content_html):
    return (
        '<div style="background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%);'
        'border-radius:16px;padding:28px 32px;margin-bottom:20px;">'
        + content_html
        + "</div>"
    )


def stat_card(label, value, accent="#D4D8C8"):
    return (
        '<div style="background:#161616;border:1px solid #1E1E1E;border-radius:16px;'
        'padding:22px;text-align:center;">'
        f'<p style="color:#888;font-size:13px;margin:0 0 6px 0;font-weight:500;">{label}</p>'
        f'<p style="color:{accent};font-size:24px;font-weight:700;margin:0;">{value}</p>'
        "</div>"
    )


# ===================================================================
#  SAYFA: Giris / Kayit
# ===================================================================
def render_auth_page():
    st.markdown(
'<div style="text-align:center;padding:40px 0 10px 0;">'
'<h1 style="color:#F5F5F7;font-size:42px;font-weight:800;letter-spacing:-1px;margin:0;">Finanser</h1>'
'<p style="color:#888;font-size:15px;margin:8px 0 0 0;">On Muhasebe & Finansal Takip</p>'
'</div>', unsafe_allow_html=True)

    if "auth_mode" not in st.session_state:
        st.session_state.auth_mode = "login"

    col_l, col_m, col_r = st.columns([1, 1.4, 1])
    with col_m:
        if st.session_state.auth_mode == "login":
            render_login_form()
        else:
            render_register_form()


def render_login_form():
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:700;">Giris Yap</h3>'
        '<p style="color:#666;font-size:13px;margin:0 0 16px 0;">Hesabiniza erisim saglayiniz</p>'
    ), unsafe_allow_html=True)

    with st.form("login_form", clear_on_submit=False):
        email = st.text_input("E-posta", placeholder="ornek@mail.com", key="login_email")
        pw = st.text_input("Sifre", type="password", placeholder="Sifrenizi giriniz", key="login_pw")
        submitted = st.form_submit_button("Giris Yap", use_container_width=True)
        if submitted:
            if not email or not pw:
                st.error("Lutfen tum alanlari doldurunuz.")
            else:
                res = login_user(email, pw)
                if res["ok"]:
                    st.query_params["token"] = res["token"]
                    st.rerun()
                else:
                    st.error(res["err"])

    st.markdown('<p style="text-align:center;margin-top:12px;color:#888;font-size:13px;">'
                'Hesabiniz yok mu?</p>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        if st.button("Kayit Ol", key="go_register", use_container_width=True):
            st.session_state.auth_mode = "register"
            st.rerun()


def render_register_form():
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:700;">Kayit Ol</h3>'
        '<p style="color:#666;font-size:13px;margin:0 0 16px 0;">Yeni hesap olusturunuz</p>'
    ), unsafe_allow_html=True)

    with st.form("register_form", clear_on_submit=False):
        ad = st.text_input("Ad Soyad", placeholder="Adiniz Soyadiniz", key="reg_ad")
        email = st.text_input("E-posta", placeholder="ornek@mail.com", key="reg_email")
        pw = st.text_input("Sifre", type="password", placeholder="En az 6 karakter", key="reg_pw")
        pw2 = st.text_input("Sifre Tekrar", type="password", placeholder="Sifrenizi tekrar giriniz", key="reg_pw2")
        ref = st.text_input("Referans Kodu (Istege Bagli)", placeholder="Ornek: Finans101", key="reg_ref")
        submitted = st.form_submit_button("Kayit Ol", use_container_width=True)
        if submitted:
            if not ad or not email or not pw or not pw2:
                st.error("Lutfen tum zorunlu alanlari doldurunuz.")
            elif len(pw) < 6:
                st.error("Sifre en az 6 karakter olmalidir.")
            elif pw != pw2:
                st.error("Sifreler eslesmiyor.")
            elif ref and ref.strip().lower() not in VALID_REF_CODES:
                st.error("Gecersiz referans kodu. Gecerli kodlar: Finans101, Finans102, finans103")
            else:
                res = register_user(ad.strip(), email.strip(), pw, ref.strip() if ref else "")
                if res["ok"]:
                    st.query_params["token"] = res["token"]
                    st.rerun()
                else:
                    st.error(res["err"])

    st.markdown('<p style="text-align:center;margin-top:12px;color:#888;font-size:13px;">'
                'Zaten hesabiniz var mi?</p>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        if st.button("Giris Yap", key="go_login", use_container_width=True):
            st.session_state.auth_mode = "login"
            st.rerun()


# ===================================================================
#  SAYFA: Genel Bakis (Dashboard)
# ===================================================================
def render_dashboard(user):
    greeting = get_greeting(user["ad"])
    st.markdown(card_gradient(
        f'<h2 style="color:#0D0D0D;margin:0;font-weight:800;font-size:28px;">{greeting}</h2>'
        f'<p style="color:#333;margin:6px 0 0 0;font-size:14px;">Finansal durumunuzun ozeti asagida yer almaktadir.</p>'
    ), unsafe_allow_html=True)

    stats = get_user_stats(user["id"])
    gelir = stats["toplam_gelir"]
    gider = stats["toplam_gider"]
    bakiye = gelir - gider
    islem = stats["islem_sayisi"]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(stat_card("Toplam Gelir", format_tl(gelir), "#7ECF8A"), unsafe_allow_html=True)
    with c2:
        st.markdown(stat_card("Toplam Gider", format_tl(gider), "#E57373"), unsafe_allow_html=True)
    with c3:
        color = "#7ECF8A" if bakiye >= 0 else "#E57373"
        st.markdown(stat_card("Bakiye", format_tl(bakiye), color), unsafe_allow_html=True)
    with c4:
        st.markdown(stat_card("Belge Kullanimi",
                              f'{user["kullanilan_belge"]} / {user["kota"]}', "#D4D8C8"),
                    unsafe_allow_html=True)

    st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 16px 0;font-weight:700;font-size:18px;">Son Islemler</h3>'
    ), unsafe_allow_html=True)

    txns = get_user_transactions(user["id"])
    if not txns:
        st.markdown('<p style="color:#666;text-align:center;padding:20px;">Henuz islem bulunmuyor.</p>',
                    unsafe_allow_html=True)
    else:
        for t in txns[:10]:
            tip_color = "#7ECF8A" if t["islem_tipi"] == "Gelir" else "#E57373"
            tip_sign = "+" if t["islem_tipi"] == "Gelir" else "-"
            row_html = (
                '<div style="background:#161616;border:1px solid #1E1E1E;border-radius:12px;'
                'padding:14px 20px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center;">'
                '<div>'
                f'<span style="color:#F5F5F7;font-weight:600;font-size:14px;">{t.get("firma","—")}</span>'
                f'<span style="color:#555;font-size:12px;margin-left:12px;">{t.get("kategori","")}</span>'
                f'<span style="color:#444;font-size:12px;margin-left:12px;">{format_date_tr(t.get("tarih",""))}</span>'
                '</div>'
                f'<span style="color:{tip_color};font-weight:700;font-size:15px;">{tip_sign} {format_tl(t.get("tutar",0))}</span>'
                '</div>'
            )
            st.markdown(row_html, unsafe_allow_html=True)


# ===================================================================
#  SAYFA: Belge Yukle (AI ile Okuma)
# ===================================================================
def render_belge_yukle(user):
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:700;">Belge Yukle</h3>'
        '<p style="color:#666;font-size:13px;margin:0;">Fatura, fis, dekont veya ekstre yukleyerek yapay zeka ile otomatik okutun.</p>'
    ), unsafe_allow_html=True)

    user = get_user_info(user["id"])
    if user["kullanilan_belge"] >= user["kota"]:
        st.markdown(card_gradient(
            '<h3 style="color:#0D0D0D;margin:0 0 8px 0;font-weight:700;">Belge Kotaniz Doldu</h3>'
            f'<p style="color:#333;margin:0;font-size:14px;">Mevcut paketiniz: {user["paket"]} — '
            f'{user["kullanilan_belge"]}/{user["kota"]} belge kullanildi.</p>'
            '<p style="color:#333;margin:8px 0 0 0;font-size:14px;">'
            'Belge yuklemeye devam etmek icin paketinizi yukseltiniz.</p>'
        ), unsafe_allow_html=True)
        if st.button("Paket Yukselt", key="quota_upgrade"):
            st.session_state.nav_page = "Paket Yukselt"
            st.rerun()
        return

    st.markdown(
        f'<p style="color:#888;font-size:13px;margin-bottom:12px;">Kalan belge hakki: '
        f'{user["kota"] - user["kullanilan_belge"]} / {user["kota"]}</p>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Belge secin (JPG, PNG veya PDF)",
        type=["jpg", "jpeg", "png", "pdf"],
        key="doc_uploader",
    )

    if uploaded:
        if uploaded.type and uploaded.type.startswith("image"):
            st.image(uploaded, caption="Yuklenen belge", use_container_width=True)

        if st.button("Belgeyi Oku", key="read_doc"):
            with st.spinner("Belge yapay zeka ile okunuyor..."):
                file_bytes = uploaded.read()
                mime = uploaded.type or "application/octet-stream"
                result = process_document_with_ai(file_bytes, mime)
                if result["ok"]:
                    st.session_state.ai_result = result["data"]
                    st.rerun()
                else:
                    st.error(f"Belge okunurken hata olustu: {result['err']}")

    if "ai_result" in st.session_state:
        data = st.session_state.ai_result
        st.markdown(card_gradient(
            '<h3 style="color:#0D0D0D;margin:0 0 4px 0;font-weight:700;">Okunan Veriler</h3>'
            '<p style="color:#333;font-size:13px;margin:0;">Asagidaki bilgileri kontrol edip gerekirse duzenleyiniz.</p>'
        ), unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            ai_tarih = st.text_input("Tarih", value=data.get("tarih", ""), key="ai_tarih")
            ai_firma = st.text_input("Firma / Kisi", value=data.get("firma", ""), key="ai_firma")
            tipler = ["Gelir", "Gider"]
            ai_idx = 0 if data.get("islem_tipi", "Gider") == "Gelir" else 1
            ai_tip = st.selectbox("Islem Tipi", tipler, index=ai_idx, key="ai_tip")
            ai_kategori = st.text_input("Kategori", value=data.get("kategori", ""), key="ai_kat")
        with c2:
            ai_tutar = st.number_input("Tutar (TL)", value=float(data.get("tutar", 0)), min_value=0.0,
                                       format="%.2f", key="ai_tutar")
            ai_kdv = st.number_input("KDV (TL)", value=float(data.get("kdv", 0)), min_value=0.0,
                                     format="%.2f", key="ai_kdv")
            ai_belge = st.text_input("Belge Turu", value=data.get("belge_turu", ""), key="ai_belge")
            ai_aciklama = st.text_input("Aciklama", value=data.get("aciklama", ""), key="ai_aciklama")

        bc1, bc2, bc3 = st.columns([1, 1, 1])
        with bc1:
            if st.button("Onayla ve Kaydet", key="ai_save"):
                add_transaction(
                    user["id"], ai_tarih, ai_tip, ai_kategori, ai_firma,
                    ai_tutar, ai_kdv, ai_aciklama, ai_belge, "Yapay Zeka",
                )
                increment_belge(user["id"])
                del st.session_state.ai_result
                st.success("Islem basariyla kaydedildi.")
                st.rerun()
        with bc2:
            if st.button("Iptal", key="ai_cancel"):
                del st.session_state.ai_result
                st.rerun()


# ===================================================================
#  SAYFA: Manuel Giris
# ===================================================================
def render_manuel_giris(user):
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:700;">Manuel Veri Girisi</h3>'
        '<p style="color:#666;font-size:13px;margin:0;">Gelir veya gider bilgilerinizi elle giriniz.</p>'
    ), unsafe_allow_html=True)

    with st.form("manual_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            m_tarih = st.date_input("Tarih", value=date.today(), format="DD.MM.YYYY", key="m_tarih")
            m_tip = st.selectbox("Islem Tipi", ["Gelir", "Gider"], key="m_tip")
            if m_tip == "Gelir":
                m_kat = st.selectbox("Kategori", GELIR_KATEGORILERI, key="m_kat")
            else:
                m_kat = st.selectbox("Kategori", GIDER_KATEGORILERI, key="m_kat")
            m_firma = st.text_input("Firma / Kisi", placeholder="Firma veya kisi adi", key="m_firma")
        with c2:
            m_tutar = st.number_input("Tutar (TL)", min_value=0.0, format="%.2f", key="m_tutar")
            m_kdv = st.number_input("KDV (TL)", min_value=0.0, format="%.2f", key="m_kdv")
            m_aciklama = st.text_input("Aciklama", placeholder="Islem hakkinda kisa not", key="m_aciklama")

        submitted = st.form_submit_button("Kaydet", use_container_width=True)
        if submitted:
            if m_tutar <= 0:
                st.error("Tutar sifirdan buyuk olmalidir.")
            elif not m_firma:
                st.error("Firma / Kisi alani bos birakilamaz.")
            else:
                tarih_str = m_tarih.strftime("%d.%m.%Y")
                add_transaction(
                    user["id"], tarih_str, m_tip, m_kat, m_firma.strip(),
                    m_tutar, m_kdv, m_aciklama.strip(), "Manuel", "Manuel",
                )
                st.success("Islem basariyla kaydedildi.")
                st.rerun()


# ===================================================================
#  SAYFA: Islemlerim
# ===================================================================
def render_islemlerim(user):
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:700;">Islemlerim</h3>'
        '<p style="color:#666;font-size:13px;margin:0;">Tum gelir ve gider kayitlariniz.</p>'
    ), unsafe_allow_html=True)

    txns = get_user_transactions(user["id"])
    if not txns:
        st.markdown('<p style="color:#666;text-align:center;padding:40px;">Henuz islem bulunmuyor.</p>',
                    unsafe_allow_html=True)
        return

    excel_data = generate_excel(txns)
    st.download_button(
        "Excel Olarak Indir (.xlsx)",
        data=excel_data,
        file_name=f"finanser_islemler_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_excel",
    )

    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)

    for t in txns:
        tip_color = "#7ECF8A" if t["islem_tipi"] == "Gelir" else "#E57373"
        tip_sign = "+" if t["islem_tipi"] == "Gelir" else "-"
        tip_label = "GELIR" if t["islem_tipi"] == "Gelir" else "GIDER"
        kdv_str = f" (KDV: {format_tl(t.get('kdv', 0))})" if t.get("kdv", 0) > 0 else ""
        kaynak_badge = ""
        if t.get("kaynak") == "Yapay Zeka":
            kaynak_badge = '<span style="background:#2A2A2A;color:#D4D8C8;font-size:10px;padding:2px 8px;border-radius:6px;margin-left:8px;">AI</span>'
        row_html = (
            '<div style="background:#161616;border:1px solid #1E1E1E;border-radius:12px;'
            'padding:16px 20px;margin-bottom:8px;">'
            '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">'
            '<div>'
            f'<span style="color:{tip_color};font-size:11px;font-weight:700;letter-spacing:.5px;">{tip_label}</span>'
            f'{kaynak_badge}'
            f'<span style="color:#555;font-size:12px;margin-left:10px;">{format_date_tr(t.get("tarih",""))}</span>'
            '</div>'
            f'<span style="color:{tip_color};font-weight:700;font-size:16px;">{tip_sign} {format_tl(t.get("tutar",0))}{kdv_str}</span>'
            '</div>'
            '<div>'
            f'<span style="color:#F5F5F7;font-weight:600;font-size:14px;">{t.get("firma","—")}</span>'
            f'<span style="color:#555;font-size:12px;margin-left:12px;">{t.get("kategori","")}</span>'
            '</div>'
        )
        if t.get("aciklama"):
            row_html += f'<p style="color:#666;font-size:12px;margin:4px 0 0 0;">{t["aciklama"]}</p>'
        row_html += "</div>"
        st.markdown(row_html, unsafe_allow_html=True)


# ===================================================================
#  SAYFA: Paket Yukselt
# ===================================================================
def render_paket_yukselt(user):
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:700;">Paket Yukselt</h3>'
        '<p style="color:#666;font-size:13px;margin:0;">Mevcut paketinizi yukselterek belge kotanizi artirin.</p>'
    ), unsafe_allow_html=True)

    current_html = (
        '<div style="background:#1A1A1A;border:1px solid #2A2A2A;border-radius:12px;'
        'padding:16px 20px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:center;">'
        '<div>'
        '<span style="color:#888;font-size:13px;">Mevcut Paket:</span>'
        f'<span style="color:#F5F5F7;font-weight:700;font-size:15px;margin-left:10px;">{user["paket"]}</span>'
        '</div>'
        '<div>'
        '<span style="color:#888;font-size:13px;">Kullanim:</span>'
        f'<span style="color:#D4D8C8;font-weight:700;font-size:15px;margin-left:10px;">{user["kullanilan_belge"]} / {user["kota"]}</span>'
        '</div>'
        '</div>'
    )
    st.markdown(current_html, unsafe_allow_html=True)

    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown(
'<div style="background:#161616;border:1px solid #2A2A2A;border-radius:16px;padding:32px;text-align:center;">'
'<p style="color:#D4D8C8;font-size:12px;font-weight:600;letter-spacing:1px;margin:0 0 8px 0;">AYLIK</p>'
'<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:800;font-size:22px;">Premium Paket</h3>'
'<p style="color:#888;font-size:13px;margin:0 0 20px 0;">Aylik 1.500 belge yukleme hakki</p>'
'<p style="color:#F5F5F7;font-size:32px;font-weight:800;margin:0 0 4px 0;">600 TL</p>'
'<p style="color:#666;font-size:13px;margin:0 0 24px 0;">/ ay</p>'
'<div style="text-align:left;margin-bottom:20px;">'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Aylik 1.500 fatura/fis/dekont</p>'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Yapay zeka ile otomatik okuma</p>'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Excel ile indirme</p>'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Oncelikli destek</p>'
'</div>'
'</div>', unsafe_allow_html=True)
        if st.button("Satin Al ve Yukselt", key="buy_monthly", use_container_width=True):
            if PAYTR_AYLIK_LINK:
                st.markdown(f'<meta http-equiv="refresh" content="0;url={PAYTR_AYLIK_LINK}">',
                            unsafe_allow_html=True)
            else:
                st.info("Odeme altyapisi yaklasimda. Lutfen daha sonra tekrar deneyiniz.")

    with c2:
        st.markdown(
'<div style="background:#161616;border:2px solid #D4D8C8;border-radius:16px;padding:32px;text-align:center;position:relative;">'
'<div style="position:absolute;top:-12px;left:50%;transform:translateX(-50%);background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%);'
'color:#0D0D0D;font-size:11px;font-weight:700;padding:4px 16px;border-radius:20px;letter-spacing:.5px;">AVANTAJLI</div>'
'<p style="color:#D4D8C8;font-size:12px;font-weight:600;letter-spacing:1px;margin:0 0 8px 0;">YILLIK</p>'
'<h3 style="color:#F5F5F7;margin:0 0 4px 0;font-weight:800;font-size:22px;">Premium Paket</h3>'
'<p style="color:#888;font-size:13px;margin:0 0 20px 0;">Aylik 1.500 belge yukleme hakki</p>'
'<p style="color:#F5F5F7;font-size:32px;font-weight:800;margin:0 0 4px 0;">5.500 TL</p>'
'<p style="color:#666;font-size:13px;margin:0 0 24px 0;">/ yil (aylik ~458 TL)</p>'
'<div style="text-align:left;margin-bottom:20px;">'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Aylik 1.500 fatura/fis/dekont</p>'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Yapay zeka ile otomatik okuma</p>'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Excel ile indirme</p>'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; Oncelikli destek</p>'
'<p style="color:#999;font-size:13px;margin:4px 0;">&#10003; 2 ay tasarruf</p>'
'</div>'
'</div>', unsafe_allow_html=True)
        if st.button("Satin Al ve Yukselt", key="buy_yearly", use_container_width=True):
            if PAYTR_YILLIK_LINK:
                st.markdown(f'<meta http-equiv="refresh" content="0;url={PAYTR_YILLIK_LINK}">',
                            unsafe_allow_html=True)
            else:
                st.info("Odeme altyapisi yaklasimda. Lutfen daha sonra tekrar deneyiniz.")


# ===================================================================
#  SAYFA: Yonetim Paneli (Sadece Admin)
# ===================================================================
def render_admin_panel():
    st.markdown(card_gradient(
        '<h2 style="color:#0D0D0D;margin:0;font-weight:800;font-size:24px;">Yonetim Paneli</h2>'
        '<p style="color:#333;margin:6px 0 0 0;font-size:14px;">Tum kullanici ve islem verilerini buradan yonetiniz.</p>'
    ), unsafe_allow_html=True)

    all_users = get_all_users()
    total_belge = get_total_belge()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(stat_card("Toplam Uye", str(len(all_users)), "#D4D8C8"), unsafe_allow_html=True)
    with c2:
        st.markdown(stat_card("Toplam Okutulan Belge", str(total_belge), "#D4D8C8"), unsafe_allow_html=True)

    st.markdown('<div style="height:16px;"></div>', unsafe_allow_html=True)
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 16px 0;font-weight:700;font-size:18px;">Kullanici Listesi</h3>'
    ), unsafe_allow_html=True)

    for u in all_users:
        ref_str = u.get("referans_kodu", "") or "-"
        reg_date = format_date_tr(u.get("created_at", ""))
        row_html = (
            '<div style="background:#161616;border:1px solid #1E1E1E;border-radius:12px;'
            'padding:14px 20px;margin-bottom:6px;">'
            f'<span style="color:#F5F5F7;font-weight:600;font-size:14px;">{u["ad"]} / {u["email"]}</span>'
            f'<span style="color:#555;font-size:13px;"> &mdash; </span>'
            f'<span style="color:#D4D8C8;font-weight:600;font-size:13px;">{u["paket"]}</span>'
            f'<span style="color:#555;font-size:13px;"> &mdash; </span>'
            f'<span style="color:#888;font-size:13px;">Kullanim: {u["kullanilan_belge"]}/{u["kota"]}</span>'
            f'<span style="color:#555;font-size:13px;"> &mdash; </span>'
            f'<span style="color:#888;font-size:13px;">Referans: {ref_str}</span>'
            f'<span style="color:#555;font-size:13px;"> &mdash; </span>'
            f'<span style="color:#888;font-size:13px;">Kayit: {reg_date}</span>'
            '</div>'
        )
        st.markdown(row_html, unsafe_allow_html=True)

    st.markdown('<div style="height:20px;"></div>', unsafe_allow_html=True)
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 16px 0;font-weight:700;font-size:18px;">Paket Degistir</h3>'
    ), unsafe_allow_html=True)

    user_options = {f"{u['ad']} ({u['email']})": u["id"] for u in all_users}
    if user_options:
        selected_user_label = st.selectbox("Kullanici Secin", list(user_options.keys()), key="admin_user_sel")
        selected_uid = user_options[selected_user_label]
        selected_u = get_user_info(selected_uid)
        paket_options = ["Ucretsiz", "Aylik Premium", "Yillik Premium"]
        current_idx = 0
        if selected_u["paket"] in paket_options:
            current_idx = paket_options.index(selected_u["paket"])
        new_paket = st.selectbox("Yeni Paket", paket_options, index=current_idx, key="admin_paket_sel")
        if st.button("Paketi Guncelle", key="admin_update_pkg"):
            update_user_package(selected_uid, new_paket)
            st.success(f"{selected_user_label} kullanicisinin paketi '{new_paket}' olarak guncellendi.")
            st.rerun()

    st.markdown('<div style="height:20px;"></div>', unsafe_allow_html=True)
    st.markdown(card_dark(
        '<h3 style="color:#F5F5F7;margin:0 0 16px 0;font-weight:700;font-size:18px;">Son Islemler (Tum Kullanicilar)</h3>'
    ), unsafe_allow_html=True)

    all_txns = get_all_transactions_admin()
    if not all_txns:
        st.markdown('<p style="color:#666;text-align:center;padding:20px;">Henuz islem bulunmuyor.</p>',
                    unsafe_allow_html=True)
    else:
        for t in all_txns[:50]:
            tip_color = "#7ECF8A" if t.get("islem_tipi") == "Gelir" else "#E57373"
            tip_sign = "+" if t.get("islem_tipi") == "Gelir" else "-"
            row_html = (
                '<div style="background:#161616;border:1px solid #1E1E1E;border-radius:12px;'
                'padding:12px 20px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center;">'
                '<div>'
                f'<span style="color:#D4D8C8;font-size:12px;font-weight:600;">{t.get("ad","?")} / {t.get("email","?")}</span>'
                f'<span style="color:#444;font-size:12px;margin-left:10px;">{format_date_tr(t.get("tarih",""))}</span>'
                f'<span style="color:#555;font-size:12px;margin-left:10px;">{t.get("firma","—")}</span>'
                f'<span style="color:#555;font-size:12px;margin-left:10px;">{t.get("kategori","")}</span>'
                '</div>'
                f'<span style="color:{tip_color};font-weight:700;font-size:14px;">{tip_sign} {format_tl(t.get("tutar",0))}</span>'
                '</div>'
            )
            st.markdown(row_html, unsafe_allow_html=True)


# ===================================================================
#  ANA UYGULAMA MANTIGI
# ===================================================================
def main():
    init_db()
    inject_css()

    token = st.query_params.get("token")
    user = get_user_by_token(token) if token else None

    if not user:
        render_auth_page()
        return

    user = get_user_info(user["id"])
    if not user:
        if "token" in st.query_params:
            del st.query_params["token"]
        render_auth_page()
        return

    # --- Sidebar ---
    with st.sidebar:
        st.markdown(
'<div style="padding:16px 0 8px 0;text-align:center;">'
'<h2 style="color:#F5F5F7;font-weight:800;font-size:26px;letter-spacing:-0.5px;margin:0;">Finanser</h2>'
'<div style="width:40px;height:3px;background:linear-gradient(135deg,#E8E6D9 0%,#D4D8C8 50%,#C5C6C7 100%);'
'margin:8px auto 0 auto;border-radius:2px;"></div>'
'</div>', unsafe_allow_html=True)

        st.markdown(
'<div style="background:#1A1A1A;border-radius:12px;padding:14px 16px;margin:16px 0;">'
f'<p style="color:#F5F5F7;font-weight:600;font-size:14px;margin:0;">{user["ad"]}</p>'
f'<p style="color:#666;font-size:12px;margin:4px 0 0 0;">{user["email"]}</p>'
f'<p style="color:#D4D8C8;font-size:11px;font-weight:600;margin:6px 0 0 0;">'
f'{user["paket"]} &mdash; {user["kullanilan_belge"]}/{user["kota"]} belge</p>'
'</div>', unsafe_allow_html=True)

        pages = ["Genel Bakis", "Belge Yukle", "Manuel Giris", "Islemlerim", "Paket Yukselt"]
        if user["rol"] == "admin":
            pages.append("Yonetim Paneli")

        default_idx = 0
        if "nav_page" in st.session_state and st.session_state.nav_page in pages:
            default_idx = pages.index(st.session_state.nav_page)

        page = st.radio("Menu", pages, index=default_idx, label_visibility="collapsed", key="sidebar_nav")
        st.session_state.nav_page = page

        st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)
        if st.button("Cikis Yap", key="logout_btn", use_container_width=True):
            logout_user(user["id"])
            if "token" in st.query_params:
                del st.query_params["token"]
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    # --- Sayfa Yonlendirme ---
    if page == "Genel Bakis":
        render_dashboard(user)
    elif page == "Belge Yukle":
        render_belge_yukle(user)
    elif page == "Manuel Giris":
        render_manuel_giris(user)
    elif page == "Islemlerim":
        render_islemlerim(user)
    elif page == "Paket Yukselt":
        render_paket_yukselt(user)
    elif page == "Yonetim Paneli" and user["rol"] == "admin":
        render_admin_panel()


main()
