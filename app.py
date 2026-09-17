import base64
from datetime import datetime
import io
from io import BytesIO
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import qrcode
import requests
import json
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

FILE_PATH = "objednavky_deni_b2b.csv"
SETTINGS_PATH = "nastaveni_deni.json"
LOGO_PATH = "logo.jpg" 
BANK_ACCOUNT = "123456789"  # Doplňte číslo účtu Deni Candle
BANK_CODE = "0800"       # Doplňte kód banky

GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO = st.secrets.get("GITHUB_REPO", "")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "deni2026")
VÝCHOZÍ_B2B_PIN = st.secrets.get("B2B_PIN", "partner2026")

EMAIL_SENDER = st.secrets.get("EMAIL_SENDER", "")
EMAIL_PASSWORD = st.secrets.get("EMAIL_PASSWORD", "")
EMAIL_RECEIVER = st.secrets.get("EMAIL_RECEIVER", "")
SMTP_SERVER = st.secrets.get("SMTP_SERVER", "smtp.centrum.cz")
SMTP_PORT = st.secrets.get("SMTP_PORT", 465)

DOPRAVA_MOZNOSTI = {
    "📦 Zásilkovna": 120,
    "🚚 Kurýr DPD": 150,
    "🏬 Osobní odběr v dílně": 0
}
DOPRAVA_ZDARMA_OD = 5000

# Katalog produktů s podporou fotek (možno doplnit cestu např. "foto/dynulka.jpg" nebo URL)
PRODUKTY_KATALOG = [
    {"nazev": "🎃 Dýňulka", "cena_mo": 269.0, "foto": ""},
    {"nazev": "☁️ Podzimní obláček", "cena_mo": 359.0, "foto": ""},
    {"nazev": "🎃 Dýňový okamžik", "cena_mo": 359.0, "foto": ""},
    {"nazev": "⛄ Dýňulka sněhulka", "cena_mo": 269.0, "foto": ""},
    {"nazev": "🍋 Citronela", "cena_mo": 259.9, "foto": ""},
    {"nazev": "❤️ Děkuji!", "cena_mo": 259.0, "foto": ""},
    {"nazev": "🌸 Chvíle pro tebe", "cena_mo": 259.0, "foto": ""},
    {"nazev": "🌺 Svítím pro tebe", "cena_mo": 259.0, "foto": ""},
    {"nazev": "🤍 Pro radost…", "cena_mo": 259.0, "foto": ""},
    {"nazev": "🕯️ Rozsviť si den", "cena_mo": 259.0, "foto": ""},
    {"nazev": "✨ Jen tak…", "cena_mo": 259.0, "foto": ""},
    {"nazev": "🌙 Vypni svět, zapal svíčku", "cena_mo": 259.0, "foto": ""},
    {"nazev": "🌼 Nebeská kopretina", "cena_mo": 249.0, "foto": ""},
    {"nazev": "💜 Nebe na dlani", "cena_mo": 249.0, "foto": ""},
    {"nazev": "👑 Královská perla – Perleť", "cena_mo": 349.0, "foto": ""},
    {"nazev": "🌸 Královská perla – růžová", "cena_mo": 349.0, "foto": ""},
    {"nazev": "🟡 Královská perla – Zlatá", "cena_mo": 349.0, "foto": ""},
    {"nazev": "💜 Pastelová elegance – frézie", "cena_mo": 249.9, "foto": ""},
    {"nazev": "🌷 Pastelová elegance", "cena_mo": 249.9, "foto": ""},
    {"nazev": "💜 Jarní pohlazení", "cena_mo": 249.9, "foto": ""},
    {"nazev": "🌼 Žlutý květ", "cena_mo": 249.9, "foto": ""},
    {"nazev": "🌹 Růžový květ", "cena_mo": 249.9, "foto": ""},
    {"nazev": "💜 Fialový květ", "cena_mo": 249.9, "foto": ""},
    {"nazev": "❤️ Srdíčko z lásky", "cena_mo": 249.9, "foto": ""}
]

st.set_page_config(page_title="Deni Candle | B2B Velkoobchod", layout="wide", page_icon="🕯️")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    [data-testid="stHeaderActionElements"] {visibility: hidden;}
    header {background-color: transparent !important;}
    footer {visibility: hidden;}
    
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stSidebar"] {
        background-color: #FAF4EE !important;
    }
    html, body, [data-testid="stAppViewContainer"] *, [data-testid="stSidebar"] * {
        color: #1A1A1A !important;
    }
    h1, h2, h3, h4 {
        color: #5C3A2E !important;
        font-family: 'Georgia', serif;
    }
    div[data-testid="stColumn"] {
        background: #F4EBE2 !important;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #E2D3C4 !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.03);
    }
    
    div[data-testid="stNumberInput"] div[data-baseweb="input"],
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        border: 1px solid #C8B8A8 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="input"] button,
    div[data-baseweb="input"] div {
        background-color: #EADCD0 !important;
        border: none !important;
    }
    div[data-baseweb="input"] button svg {
        fill: #1A1A1A !important;
        color: #1A1A1A !important;
    }
    div[data-testid="stNumberInput"] input {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
    }
    div[data-testid="stNumberInput"] button {
        background-color: #EADCD0 !important;
        border: none !important;
        color: #1A1A1A !important;
    }
    div[data-testid="stNumberInput"] button:hover {
        background-color: #D8C6B6 !important;
    }

    code {
        background-color: #EADCD0 !important;
        color: #1A1A1A !important;
        border: 1px solid #C8B8A8 !important;
        padding: 2px 6px !important;
        border-radius: 4px !important;
        font-weight: bold !important;
    }

    input, textarea {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
        border: 1px solid #C8B8A8 !important;
        border-radius: 6px !important;
    }

    div.stButton > button,
    div.stDownloadButton > button,
    div.stFormSubmitButton > button,
    [data-testid="stDownloadButton"] > button,
    [data-testid="stFormSubmitButton"] > button {
        background-color: #8C5A47 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: bold !important;
    }
    div.stButton > button *,
    div.stDownloadButton > button *,
    div.stFormSubmitButton > button *,
    [data-testid="stDownloadButton"] > button *,
    [data-testid="stFormSubmitButton"] > button * {
        color: #FFFFFF !important;
    }
    div.stButton > button:hover,
    div.stDownloadButton > button:hover,
    div.stFormSubmitButton > button:hover {
        background-color: #6E4434 !important;
    }

    div[data-testid="stTable"], 
    div[data-testid="stTable"] table {
        background-color: #FAF4EE !important;
        color: #1A1A1A !important;
        width: 100% !important;
    }
    table {
        background-color: #FAF4EE !important;
        color: #1A1A1A !important;
        border-collapse: collapse !important;
    }
    th {
        background-color: #EADCD0 !important;
        color: #5C3A2E !important;
        font-weight: bold !important;
        border-bottom: 2px solid #C8B8A8 !important;
        padding: 10px !important;
    }
    td {
        background-color: #FAF4EE !important;
        color: #1A1A1A !important;
        border-bottom: 1px solid #E2D3C4 !important;
        padding: 8px 10px !important;
    }

    .summary-card {
        background-color: #EADCD0 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #8C5A47;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

def get_headers():
    return {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}

def odeslat_email_upozorneni(id_obj, firma, cena, doprava_nazev, celkem_ks, polozky_text):
    if not EMAIL_SENDER or not EMAIL_PASSWORD or not EMAIL_RECEIVER:
        return False
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Subject'] = f"🕯️ Nová B2B objednávka #{id_obj} od {firma}"
        
        body = f"""Dobrý den,

přes B2B portál Deni Candle byla právě přijata nová objednávka!

🛍️ Objednávka #{id_obj}
🏢 Partner: {firma}
🚚 Doprava: {doprava_nazev}
💰 Celková cena (VO): {cena:,.0f} Kč
📦 Celkem kusů: {celkem_ks} ks

Přehled položek:
{polozky_text}

Objednávku si můžete detailně prohlédnout v administraci aplikace.

Hezký den,
Váš B2B systém Deni Candle
"""
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        if SMTP_PORT == 465:
            server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        else:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Chyba pri odesilani e-mailu: {e}")
        return False

def nacti_nastaveni():
    vychozi_data = {
        "partneri": [
            {"PIN": VÝCHOZÍ_B2B_PIN, "Nazev": "Základní velkoodběratel", "Sleva": 40}
        ]
    }
    if not GITHUB_TOKEN or not GITHUB_REPO:
        if not os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "w") as f:
                json.dump(vychozi_data, f)
        with open(SETTINGS_PATH, "r") as f:
            return json.load(f), None

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{SETTINGS_PATH}"
    res = requests.get(url, headers=get_headers())
    if res.status_code == 200:
        data = res.json()
        sha = data["sha"]
        content_str = base64.b64decode(data["content"]).decode("utf-8")
        return json.loads(content_str), sha
    else:
        return vychozi_data, None

def uloz_nastaveni(nastaveni_dict, sha=None):
    obsah = json.dumps(nastaveni_dict)
    if not GITHUB_TOKEN or not GITHUB_REPO:
        with open(SETTINGS_PATH, "w") as f:
            f.write(obsah)
        return True

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{SETTINGS_PATH}"
    content_b64 = base64.b64encode(obsah.encode("utf-8")).decode("utf-8")
    payload = {"message": "Aktualizace slev Deni Candle", "content": content_b64}
    if sha:
        payload["sha"] = sha
    res = requests.put(url, headers=get_headers(), json=payload)
    return res.status_code in [200, 201]

def nacti_objednavky():
    sloupce = [
        "ID", "Datum_Vytvoreni", "Oznaceni_Partnera", "Firma_ICO", "Jmeno_Kontakt", 
        "Telefon", "Email", "Adresa_Doruceni", "Poznamka", "Polozky_Detail", 
        "Celkem_Ks", "Sleva_Pouzita", "Cena_Zbozi_VO", "Doprava_Nazev", "Cena_Dopravy", 
        "Cena_Celkem_VO", "Stav_Platby", "Stav_Vyroby"
    ]
    
    if not GITHUB_TOKEN or not GITHUB_REPO:
        if not os.path.exists(FILE_PATH):
            df_empty = pd.DataFrame(columns=sloupce)
            df_empty.to_csv(FILE_PATH, index=False)
            return df_empty, None
        df = pd.read_csv(FILE_PATH)
        for col in ["Stav_Vyroby", "Doprava_Nazev", "Cena_Zbozi_VO"]:
            if col not in df.columns:
                df[col] = "K výrobě" if col == "Stav_Vyroby" else ("Osobní odběr" if col == "Doprava_Nazev" else df.get("Cena_Celkem_VO", 0))
        return df, None

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}"
    res = requests.get(url, headers=get_headers())
    if res.status_code == 200:
        data = res.json()
        sha = data["sha"]
        content_str = base64.b64decode(data["content"]).decode("utf-8")
        df = pd.read_csv(io.StringIO(content_str))
        for col in ["Stav_Vyroby", "Doprava_Nazev", "Cena_Zbozi_VO"]:
            if col not in df.columns:
                df[col] = "K výrobě" if col == "Stav_Vyroby" else ("Osobní odběr" if col == "Doprava_Nazev" else df.get("Cena_Celkem_VO", 0))
        return df, sha
    else:
        df_empty = pd.DataFrame(columns=sloupce)
        return df_empty, None

def uloz_objednavky(df, sha=None):
    if not GITHUB_TOKEN or not GITHUB_REPO:
        df.to_csv(FILE_PATH, index=False)
        return True

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}"
    csv_buffer = io.StringIO()
    df.to_csv(csv_buffer, index=False)
    content_b64 = base64.b64encode(csv_buffer.getvalue().encode("utf-8")).decode("utf-8")

    payload = {"message": "Aktualizace B2B objednavek Deni Candle", "content": content_b64}
    if sha:
        payload["sha"] = sha
    res = requests.put(url, headers=get_headers(), json=payload)
    return res.status_code in [200, 201]

def vytvor_profi_excel(df, titulek="Objednávky"):
    wb = Workbook()
    ws = wb.active
    ws.title = titulek
    
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
        
    hlavicka_fill = PatternFill(start_color="8C5A47", end_color="8C5A47", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )
    
    for cell in ws[1]:
        cell.fill = hlavicka_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[col_letter].width = (max_length + 2)
        
    ws.auto_filter.ref = ws.dimensions
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

def vygeneruj_b2b_uctenku(id_obj, firma, jmeno, adresa, telefon, polozky_str, celkem_ks, cena_zbozi, doprava_nazev, cena_dopravy, cena_celkem, sleva):
    return f"""
    <!DOCTYPE html>
    <html lang="cs">
    <head>
    <meta charset="UTF-8">
    <title>B2B Objednávka Deni Candle #{id_obj}</title>
    <style>
        body {{ font-family: Arial, sans-serif; color: #111; max-width: 650px; margin: 0 auto; padding: 30px; background: #FAF4EE; }}
        .box {{ background: #fff; padding: 30px; border-radius: 10px; border-top: 6px solid #8C5A47; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
        .header {{ text-align: center; border-bottom: 2px solid #eee; padding-bottom: 15px; }}
        .header h1 {{ color: #8C5A47; margin: 0; }}
        .total {{ background: #F4EBE6; padding: 15px; text-align: center; font-size: 20px; font-weight: bold; color: #8C5A47; margin-top: 20px; border-radius: 6px; }}
    </style>
    </head>
    <body>
        <div class="box">
            <div class="header">
                <h1>🕯️ Deni Candle — B2B Potvrzení</h1>
                <p>Velkoobchodní objednávka #{id_obj} (Sleva {sleva} %)</p>
            </div>
            <h3>Odběratel:</h3>
            <p><strong>Firma / IČO:</strong> {firma}<br><strong>Kontakt:</strong> {jmeno} ({telefon})<br><strong>Adresa doručení:</strong> {adresa}</p>
            <h3>Objednané zboží ({celkem_ks} ks):</h3>
            <p>{polozky_str}</p>
            <p><strong>Doprava:</strong> {doprava_nazev} ({cena_dopravy:,.0f} Kč)</p>
            <div class="total">Celková cena k úhradě: {cena_celkem:,.0f} Kč</div>
            <p style="margin-top:20px; font-size:12px; color:#666; text-align:center;">Děkujeme za váš odběr rukodělných svíček Deni Candle.</p>
        </div>
    </body>
    </html>
    """

df_orders, current_sha = nacti_objednavky()
nastaveni_app, sha_nastaveni = nacti_nastaveni()

partneri_seznam = nastaveni_app.get("partneri", [{"PIN": VÝCHOZÍ_B2B_PIN, "Nazev": "Základní velkoodběratel", "Sleva": 40}])
partneri_dict = {str(p["PIN"]): p for p in partneri_seznam}

if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, width=140)

st.sidebar.markdown("## 🕯️ Deni Candle B2B")
rezim = st.sidebar.radio("Navigace:", ["🛍️ Velkoobchodní objednávka", "🔐 Správa pro majitele"])

# ---------------------------------------------------------
# 1. B2B OBJEDNÁVKOVÝ PORTÁL PRO PARTNERY
# ---------------------------------------------------------
if rezim == "🛍️ Velkoobchodní objednávka":
    st.markdown("<h1>🕯️ Velkoobchodní portál Deni Candle</h1>", unsafe_allow_html=True)
    
    url_pin = st.query_params.get("pin", "")
    zadan_pin = st.text_input("Zadejte přístupový B2B PIN partnera:", type="password", value=url_pin)
    
    if zadan_pin in partneri_dict:
        aktivni_partner = partneri_dict[zadan_pin]
        aktualni_sleva = aktivni_partner["Sleva"]
        nazev_partnera = aktivni_partner["Nazev"]
        
        st.success(f"✅ Vítejte, **{nazev_partnera}**! Uplatněna vaše partnerská sleva **{aktualni_sleva} %** z maloobchodních cen.")
        
        tab_p1, tab_p2 = st.tabs(["🛒 Nová objednávka", "📜 Moje historie objednávek"])
        
        with tab_p1:
            col_katalog, col_kosik = st.columns([1.3, 1])
            
            vybrane_polozky = []
            celkova_cena_zbozi_vo = 0
            celkem_ks = 0
            mo_hodnota_celkem = 0
            
            with col_katalog:
                st.subheader("1. Výběr svíček a produktů")
                col_p1, col_p2 = st.columns(2)
                
                for idx, prod in enumerate(PRODUKTY_KATALOG):
                    target_col = col_p1 if idx % 2 == 0 else col_p2
                    cena_mo = prod["cena_mo"]
                    cena_vo = round(cena_mo * (1 - aktualni_sleva / 100))
                    
                    # Pokud existuje fotka, zobrazíme ji
                    if prod.get("foto") and os.path.exists(prod["foto"]):
                        target_col.image(prod["foto"], width=100)
                    
                    label = f"{prod['nazev']} (VO: {cena_vo} Kč | MO: {cena_mo:.0f} Kč)"
                    
                    key_input = f"vo_{idx}"
                    default_val = st.session_state.get(key_input, 0)
                    ks = target_col.number_input(label, min_value=0, max_value=500, value=default_val, key=key_input)
                    
                    if ks > 0:
                        vybrane_polozky.append(f"{ks}x {prod['nazev']}")
                        celkova_cena_zbozi_vo += ks * cena_vo
                        mo_hodnota_celkem += ks * cena_mo
                        celkem_ks += ks

            with col_kosik:
                st.subheader("2. Údaje & Způsob dopravy")
                
                firma = st.text_input("Název firmy / Obchodu & IČO *", value=nazev_partnera).strip()
                jmeno = st.text_input("Jméno kontaktní osoby *").strip()
                col_t, col_e = st.columns(2)
                telefon = col_t.text_input("Telefon *").strip()
                email = col_e.text_input("E-mail *").strip()
                adresa = st.text_area("Doručovací adresa (Ulice, ČP, Město, PSČ) *", height=80).strip()
                
                st.markdown("**🚚 Výběr dopravy:**")
                zvolena_doprava_nazev = st.radio("Způsob doručení:", list(DOPRAVA_MOZNOSTI.keys()))
                zakladni_cena_dopravy = DOPRAVA_MOZNOSTI[zvolena_doprava_nazev]
                
                # Výpočet dopravy zdarma
                if celkova_cena_zbozi_vo >= DOPRAVA_ZDARMA_OD and zakladni_cena_dopravy > 0:
                    cena_dopravy = 0
                    st.success("🎉 Skvělé! Dosáhli jste na **Dopravu ZDARMA**.")
                else:
                    cena_dopravy = zakladni_cena_dopravy
                    if zakladni_cena_dopravy > 0 and celkova_cena_zbozi_vo > 0:
                        st.caption(f"💡 Doprava ZDARMA od {DOPRAVA_ZDARMA_OD:,.0f} Kč (chybí {DOPRAVA_ZDARMA_OD - celkova_cena_zbozi_vo:,.0f} Kč).")

                poznamka = st.text_input("Poznámka k doručení / balení:").strip()
                
                st.divider()
                
                celkova_konecna_cena = celkova_cena_zbozi_vo + cena_dopravy
                uspora = mo_hodnota_celkem - celkova_cena_zbozi_vo
                
                st.markdown(f"""
                    <div class='summary-card'>
                        <h4 style='margin: 0;'>Shrnutí VO objednávky</h4>
                        <p style='margin: 5px 0 0 0; font-size: 14px;'>Celkem produktů: <strong>{celkem_ks} ks</strong></p>
                        <p style='margin: 0; font-size: 13px;'>Cena zboží: {celkova_cena_zbozi_vo:,.0f} Kč</p>
                        <p style='margin: 0; font-size: 13px;'>Doprava: {cena_dopravy:,.0f} Kč</p>
                        <p style='margin: 0; font-size: 13px; color: #8C5A47;'>Úspora oproti MO: {uspora:,.0f} Kč</p>
                        <h3 style='margin: 10px 0 0 0;'>Celkem k úhradě: {celkova_konecna_cena:,.0f} Kč</h3>
                    </div>
                """, unsafe_allow_html=True)
                
                if st.button("Odeslat velkoobchodní objednávku", type="primary", use_container_width=True):
                    if not all([firma, jmeno, telefon, email, adresa]):
                        st.warning("⚠️ Prosím vyplňte všechny kontaktní a firemní údaje.")
                    elif celkem_ks == 0:
                        st.error("❌ Košík je prázdný. Vyberte prosím alespoň jednu svíčku.")
                    else:
                        with st.spinner('Odesílám VO objednávku... 🕯️'):
                            nove_id = 1 if df_orders.empty else int(df_orders["ID"].max()) + 1
                            polozky_text_email = ",\n".join(vybrane_polozky)
                            polozky_text_db = ", ".join(vybrane_polozky)
                            
                            nova_obj = pd.DataFrame([{
                                "ID": nove_id,
                                "Datum_Vytvoreni": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                "Oznaceni_Partnera": nazev_partnera,
                                "Firma_ICO": firma,
                                "Jmeno_Kontakt": jmeno,
                                "Telefon": telefon,
                                "Email": email,
                                "Adresa_Doruceni": adresa,
                                "Poznamka": poznamka,
                                "Polozky_Detail": polozky_text_db,
                                "Celkem_Ks": celkem_ks,
                                "Sleva_Pouzita": f"{aktualni_sleva} %",
                                "Cena_Zbozi_VO": celkova_cena_zbozi_vo,
                                "Doprava_Nazev": zvolena_doprava_nazev,
                                "Cena_Dopravy": cena_dopravy,
                                "Cena_Celkem_VO": celkova_konecna_cena,
                                "Stav_Platby": "Čeká na platbu",
                                "Stav_Vyroby": "K výrobě"
                            }])
                            
                            df_aktualni = pd.concat([df_orders, nova_obj], ignore_index=True)
                            if uloz_objednavky(df_aktualni, current_sha):
                                odeslat_email_upozorneni(nove_id, firma, celkova_konecna_cena, zvolena_doprava_nazev, celkem_ks, polozky_text_email)
                                
                                st.balloons()
                                st.success("🎉 Děkujeme! Velkoobchodní objednávka byla úspěšně přijata.")
                                
                                html_uct = vygeneruj_b2b_uctenku(nove_id, firma, jmeno, adresa, telefon, polozky_text_email.replace("\n", "<br>"), celkem_ks, celkova_cena_zbozi_vo, zvolena_doprava_nazev, cena_dopravy, celkova_konecna_cena, aktualni_sleva)
                                st.download_button("📥 Stáhnout B2B Potvrzení (HTML/PDF)", html_uct, file_name=f"DeniCandle_B2B_{nove_id}.html", mime="text/html")
                                
                                spd_str = f"SPD*1.0*ACC:{BANK_ACCOUNT}/{BANK_CODE}*AM:{celkova_konecna_cena:.2f}*CC:CZK*X-VS:{nove_id}*MSG:DeniCandle B2B {nove_id}"
                                qr_img = qrcode.make(spd_str)
                                buf = BytesIO()
                                qr_img.save(buf, format="PNG")
                                st.image(buf.getvalue(), caption="QR Platba převodem", width=200)
                            else:
                                st.error("❌ Chyba při ukládání objednávky.")
                                
        with tab_p2:
            st.markdown(f"### 📜 Minulé objednávky partnera **{nazev_partnera}**")
            moje_objednavky = df_orders[df_orders['Oznaceni_Partnera'] == nazev_partnera].sort_values(by="ID", ascending=False) if not df_orders.empty else pd.DataFrame()
            
            if not moje_objednavky.empty:
                for idx, row in moje_objednavky.iterrows():
                    id_o = row['ID']
                    dat_o = row['Datum_Vytvoreni']
                    cena_o = row['Cena_Celkem_VO']
                    ks_o = row['Celkem_Ks']
                    detail_o = str(row['Polozky_Detail'])
                    
                    with st.expander(f"📦 Objednávka #{id_o} — {cena_o:,.0f} Kč ({dat_o})"):
                        st.write(f"**Položky ({ks_o} ks):** {detail_o}")
                        st.write(f"**Doprava:** {row.get('Doprava_Nazev', '-')} | **Stav:** {row.get('Stav_Vyroby', 'K výrobě')}")
                        
                        if st.button(f"🔄 Zopakovat objednávku #{id_o}", key=f"repeat_{id_o}"):
                            # Načtení položek z historie do formuláře
                            for i_p, prod in enumerate(PRODUKTY_KATALOG):
                                st.session_state[f"vo_{i_p}"] = 0
                                for item in detail_o.split(", "):
                                    if "x " in item:
                                        k_str, n_str = item.split("x ", 1)
                                        if n_str.strip() == prod["nazev"].strip():
                                            st.session_state[f"vo_{i_p}"] = int(k_str)
                            st.success(f"Položky z objednávky #{id_o} byly načteny do košíku! Přepněte se na záložku 'Nová objednávka'.")
                            st.rerun()
            else:
                st.info("Zatím jste nevytvořili žádné objednávky.")

    elif zadan_pin:
        st.error("❌ Neplatný B2B PIN. Zkontrolujte prosím přístupové heslo.")

# ---------------------------------------------------------
# 2. SPRÁVA PRO MAJITELE DENI CANDLE
# ---------------------------------------------------------
else:
    st.markdown("<h1 class='main-header'>🔐 Správa Deni Candle</h1>", unsafe_allow_html=True)
    heslo_admin = st.text_input("Heslo správce:", type="password")
    
    if heslo_admin == ADMIN_PASSWORD:
        st.success("✅ Přístup schválen.")
        
        tab_db, tab_vyroba, tab_partneri = st.tabs(["📋 Databáze objednávek", "🕯️ Souhrn pro výrobu", "👥 Správa odběratelů a slev"])
        
        with tab_db:
            if not df_orders.empty:
                st.markdown("### 📋 Přehled přijatých objednávek")
                
                excel_db = vytvor_profi_excel(df_orders, titulek="Databaze_Objednavek")
                st.download_button(
                    label="📥 Stáhnout kompletní databázi do Excelu", 
                    data=excel_db, 
                    file_name=f"DeniCandle_Databaze_{datetime.now().strftime('%d_%m_%Y')}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_db_all"
                )
                
                st.divider()
                
                df_orders_sorted = df_orders.sort_values(by="ID", ascending=False)
                
                for idx, row in df_orders_sorted.iterrows():
                    id_obj = row['ID']
                    datum = row['Datum_Vytvoreni']
                    partner = row['Oznaceni_Partnera']
                    cena = row['Cena_Celkem_VO']
                    stav_v = row.get('Stav_Vyroby', 'K výrobě')
                    
                    ikona_stavu = "🔥" if stav_v == "K výrobě" else "✅"
                    
                    with st.expander(f"{ikona_stavu} Objednávka #{id_obj} — {partner} | {cena:,.0f} Kč | {datum} ({stav_v})"):
                        c1, c2 = st.columns(2)
                        
                        with c1:
                            st.markdown(f"**🏢 Odběratel:** {row['Firma_ICO']}")
                            st.markdown(f"**👤 Kontakt:** {row['Jmeno_Kontakt']}")
                            st.markdown(f"**📞 Telefon:** {row['Telefon']}")
                            st.markdown(f"**✉️ E-mail:** {row['Email']}")
                            st.markdown(f"**📍 Adresa:** {row['Adresa_Doruceni']}")
                            
                        with c2:
                            st.markdown(f"**🏷️ Použitá sleva:** {row['Sleva_Pouzita']}")
                            st.markdown(f"**🚚 Doprava:** {row.get('Doprava_Nazev', '-')} ({row.get('Cena_Dopravy', 0)} Kč)")
                            st.markdown(f"**💰 Celková cena (VO):** {cena:,.0f} Kč")
                            st.markdown(f"**📦 Celkem kusů:** {row['Celkem_Ks']} ks")
                            st.markdown(f"**📌 Stav zakázky:** `{stav_v}`")
                            
                            poznamka = row['Poznamka']
                            if pd.notna(poznamka) and str(poznamka).strip() != "":
                                st.markdown(f"**📝 Poznámka:** {poznamka}")
                        
                        st.markdown("**🛍️ Objednané položky:**")
                        st.info(row['Polozky_Detail'])
                        
                        col_b1, col_b2 = st.columns(2)
                        if stav_v == "K výrobě":
                            if col_b1.button(f"✅ Označit #{id_obj} jako vyřízenou/odlitou", key=f"done_db_{id_obj}"):
                                df_orders.loc[df_orders['ID'] == id_obj, 'Stav_Vyroby'] = 'Vyřízeno'
                                uloz_objednavky(df_orders, current_sha)
                                st.success("✅ Změněno na Vyřízeno.")
                                st.rerun()
                        else:
                            if col_b1.button(f"🔄 Vrátit #{id_obj} zpět k výrobě", key=f"reopen_db_{id_obj}"):
                                df_orders.loc[df_orders['ID'] == id_obj, 'Stav_Vyroby'] = 'K výrobě'
                                uloz_objednavky(df_orders, current_sha)
                                st.rerun()
                                
                        if col_b2.button(f"🗑️ Smazat objednávku #{id_obj}", key=f"del_db_{id_obj}"):
                            df_orders = df_orders[df_orders['ID'] != id_obj]
                            uloz_objednavky(df_orders, current_sha)
                            st.success("🗑️ Objednávka smazána.")
                            st.rerun()
            else:
                st.info("Zatím žádné B2B objednávky.")
                
        with tab_vyroba:
            df_k_vyrobe = df_orders[df_orders.get('Stav_Vyroby', 'K výrobě') == 'K výrobě'] if not df_orders.empty else pd.DataFrame()
            
            if not df_k_vyrobe.empty:
                st.markdown("### 📊 Výrobní matice (Pouze nezpracované zakázky)")
                
                matrix_rows = []
                for idx, row in df_k_vyrobe.iterrows():
                    id_obj = row['ID']
                    partner = row['Oznaceni_Partnera']
                    col_label = f"Obj #{id_obj} ({partner})"
                    
                    detail = row["Polozky_Detail"]
                    if pd.notna(detail):
                        for item in str(detail).split(", "):
                            if "x " in item:
                                ks, název = item.split("x ", 1)
                                matrix_rows.append({"Produkt": název, "Col": col_label, "Ks": int(ks)})
                
                if matrix_rows:
                    df_matrix_raw = pd.DataFrame(matrix_rows)
                    df_pivot = df_matrix_raw.pivot_table(index="Produkt", columns="Col", values="Ks", aggfunc="sum", fill_value=0)
                    
                    df_pivot["CELKEM KS"] = df_pivot.sum(axis=1)
                    df_pivot = df_pivot.sort_values(by="CELKEM KS", ascending=False).reset_index()
                    
                    df_pivot_display = df_pivot.copy()
                    for col in df_pivot_display.columns:
                        if col not in ["Produkt", "CELKEM KS"]:
                            df_pivot_display[col] = df_pivot_display[col].apply(lambda x: f"{x} ks" if x > 0 else "-")
                        elif col == "CELKEM KS":
                            df_pivot_display[col] = df_pivot_display[col].apply(lambda x: f"🔥 {x} ks")

                    st.table(df_pivot_display)
                    
                    excel_data = vytvor_profi_excel(df_pivot, titulek="Vyrobni_Matice")
                    st.download_button(
                        label="📥 Stáhnout výrobní matici do Excelu", 
                        data=excel_data, 
                        file_name=f"DeniCandle_VyrobniMatice_{datetime.now().strftime('%d_%m')}.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_all_vyroba_matrix"
                    )
                else:
                    st.info("Žádné svíčky k výrobě.")
                
                st.divider()
                
                st.markdown("### 📦 Samostatné rozpisky k odlítí")
                df_k_vyrobe_sorted = df_k_vyrobe.sort_values(by="ID", ascending=False)
                
                for idx, row in df_k_vyrobe_sorted.iterrows():
                    id_obj = row['ID']
                    partner = row['Oznaceni_Partnera']
                    datum = row['Datum_Vytvoreni']
                    
                    with st.expander(f"Rozpiska pro Objednávku #{id_obj} — {partner} ({datum})"):
                        polozky_obj = []
                        detail = row["Polozky_Detail"]
                        if pd.notna(detail):
                            for item in str(detail).split(", "):
                                if "x " in item:
                                    ks, název = item.split("x ", 1)
                                    polozky_obj.append({"Produkt": název, "Ks": int(ks)})
                        
                        if polozky_obj:
                            df_obj = pd.DataFrame(polozky_obj).groupby("Produkt").sum().reset_index()
                            st.table(df_obj)
                            
                            c_m1, c_m2 = st.columns(2)
                            excel_obj = vytvor_profi_excel(df_obj, titulek=f"Objednavka_{id_obj}")
                            c_m1.download_button(
                                label=f"📥 Excel pro objednávku #{id_obj}", 
                                data=excel_obj, 
                                file_name=f"DeniCandle_Vyroba_Obj_{id_obj}.xlsx", 
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_obj_{id_obj}"
                            )
                            
                            if c_m2.button(f"✅ Hotovo / Odlito (Vyřídit #{id_obj})", key=f"done_vyroba_{id_obj}"):
                                df_orders.loc[df_orders['ID'] == id_obj, 'Stav_Vyroby'] = 'Vyřízeno'
                                uloz_objednavky(df_orders, current_sha)
                                st.rerun()
            else:
                st.success("🎉 Skvělé! Všechny svíčky jsou odlité a žádné objednávky nečekají na výrobu.")
                
        with tab_partneri:
            st.markdown("### ➕ Přidat nového partnera")
            with st.form("form_novy_partner", clear_on_submit=True):
                col_p1, col_p2, col_p3 = st.columns([2, 3, 2])
                n_pin = col_p1.text_input("Přístupový PIN (heslo) *").strip()
                n_nazev = col_p2.text_input("Název partnera / obchodu *").strip()
                n_sleva = col_p3.number_input("Sleva v %", min_value=0, max_value=99, value=40)
                
                btn_pridat = st.form_submit_button("➕ Uložit nového partnera")
                if btn_pridat:
                    if not n_pin or not n_nazev:
                        st.warning("⚠️ Vyplňte prosím PIN i název partnera.")
                    else:
                        existujici_piny = [str(p["PIN"]) for p in partneri_seznam]
                        if n_pin in existujici_piny:
                            st.error("❌ Tento PIN už používá jiný partner. Zvolte jiný.")
                        else:
                            partneri_seznam.append({"PIN": n_pin, "Nazev": n_nazev, "Sleva": int(n_sleva)})
                            if uloz_nastaveni({"partneri": partneri_seznam}, sha_nastaveni):
                                st.success(f"✅ Partner **{n_nazev}** byl úspěšně přidán!")
                                st.rerun()
                            else:
                                st.error("❌ Chyba při ukládání.")

            st.divider()
            st.markdown("### 👥 Seznam aktivních odběratelů a slev")
            
            partneri_ke_smazani = []
            for i, p in enumerate(partneri_seznam):
                col1, col2, col3, col4 = st.columns([2, 3, 2, 1.2])
                col1.write(f"🔑 **PIN:** `{p['PIN']}`")
                col2.write(f"🏪 **Partner:** {p['Nazev']}")
                col3.write(f"🏷️ **Sleva:** {p['Sleva']} %")
                if col4.button("🗑️ Smazat", key=f"del_partner_{i}"):
                    partneri_ke_smazani.append(i)
                    
            if partneri_ke_smazani:
                for idx in sorted(partneri_ke_smazani, reverse=True):
                    partneri_seznam.pop(idx)
                if uloz_nastaveni({"partneri": partneri_seznam}, sha_nastaveni):
                    st.success("✅ Partner byl smazán.")
                    st.rerun()
