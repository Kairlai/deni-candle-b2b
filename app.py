import base64
from datetime import datetime, timedelta
import io
from io import BytesIO
import os
import random
import urllib.parse
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

MIN_OBJEDNAVKA_KC = 3000

PRODUKTY_KATALOG = [
    {"nazev": "🎃 Dýňulka", "cena_mo": 269.0},
    {"nazev": "☁️ Podzimní obláček", "cena_mo": 359.0},
    {"nazev": "🎃 Dýňový okamžik", "cena_mo": 359.0},
    {"nazev": "⛄ Dýňulka sněhulka", "cena_mo": 269.0},
    {"nazev": "🍋 Citronela", "cena_mo": 259.9},
    {"nazev": "❤️ Děkuji!", "cena_mo": 259.0},
    {"nazev": "🌸 Chvíle pro tebe", "cena_mo": 259.0},
    {"nazev": "🌺 Svítím pro tebe", "cena_mo": 259.0},
    {"nazev": "🤍 Pro radost…", "cena_mo": 259.0},
    {"nazev": "🕯️ Rozsviť si den", "cena_mo": 259.0},
    {"nazev": "✨ Jen tak…", "cena_mo": 259.0},
    {"nazev": "🌙 Vypni svět, zapal svíčku", "cena_mo": 259.0},
    {"nazev": "🌼 Nebeská kopretina", "cena_mo": 249.0},
    {"nazev": "💜 Nebe na dlani", "cena_mo": 249.0},
    {"nazev": "👑 Královská perla – Perleť", "cena_mo": 349.0},
    {"nazev": "🌸 Královská perla – růžová", "cena_mo": 349.0},
    {"nazev": "🟡 Královská perla – Zlatá", "cena_mo": 349.0},
    {"nazev": "💜 Pastelová elegance – frézie", "cena_mo": 249.9},
    {"nazev": "🌷 Pastelová elegance", "cena_mo": 249.9},
    {"nazev": "💜 Jarní pohlazení", "cena_mo": 249.9},
    {"nazev": "🌼 Žlutý květ", "cena_mo": 249.9},
    {"nazev": "🌹 Růžový květ", "cena_mo": 249.9},
    {"nazev": "💜 Fialový květ", "cena_mo": 249.9},
    {"nazev": "❤️ Srdíčko z lásky", "cena_mo": 249.9}
]

st.set_page_config(page_title="Deni Candle | B2B Velkoobchod", layout="wide", page_icon="🕯️")

# Vynucené béžové pozadí a černé písmo
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
    input, textarea, div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
        border: 1px solid #C8B8A8 !important;
        border-radius: 6px !important;
    }
    button[title="Increase value"], button[title="Decrease value"] {
        background-color: #E2D3C4 !important;
        color: #1A1A1A !important;
    }
    div.stButton > button:first-child {
        background-color: #8C5A47 !important;
        color: #FFFFFF !important;
        border: none;
        border-radius: 8px;
        font-weight: bold;
    }
    div.stButton > button:first-child * {
        color: #FFFFFF !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #6E4434 !important;
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
    if not GITHUB_TOKEN or not GITHUB_REPO:
        if not os.path.exists(FILE_PATH):
            df_empty = pd.DataFrame(columns=[
                "ID", "Datum_Vytvoreni", "Oznaceni_Partnera", "Firma_ICO", "Jmeno_Kontakt", 
                "Telefon", "Email", "Adresa_Doruceni", "Poznamka", "Polozky_Detail", 
                "Celkem_Ks", "Sleva_Pouzita", "Cena_Celkem_VO", "Stav_Platby"
            ])
            df_empty.to_csv(FILE_PATH, index=False)
        return pd.read_csv(FILE_PATH), None

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}"
    res = requests.get(url, headers=get_headers())
    if res.status_code == 200:
        data = res.json()
        sha = data["sha"]
        content_str = base64.b64decode(data["content"]).decode("utf-8")
        return pd.read_csv(io.StringIO(content_str)), sha
    else:
        df_empty = pd.DataFrame(columns=[
            "ID", "Datum_Vytvoreni", "Oznaceni_Partnera", "Firma_ICO", "Jmeno_Kontakt", 
            "Telefon", "Email", "Adresa_Doruceni", "Poznamka", "Polozky_Detail", 
            "Celkem_Ks", "Sleva_Pouzita", "Cena_Celkem_VO", "Stav_Platby"
        ])
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

def vygeneruj_b2b_uctenku(id_obj, firma, jmeno, adresa, telefon, polozky_str, celkem_ks, cena, sleva):
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
                <p>Velkoobchodní objednávka #{id_obj} (Aplikována sleva {sleva} %)</p>
            </div>
            <h3>Odběratel:</h3>
            <p><strong>Firma / IČO:</strong> {firma}<br><strong>Kontakt:</strong> {jmeno} ({telefon})<br><strong>Adresa doručení:</strong> {adresa}</p>
            <h3>Objednané zboží ({celkem_ks} ks):</h3>
            <p>{polozky_str}</p>
            <div class="total">Celková cena VO: {cena:,.0f} Kč</div>
            <p style="margin-top:20px; font-size:12px; color:#666; text-align:center;">Děkujeme za váš odběr rukodělných svíček Deni Candle.</p>
        </div>
    </body>
    </html>
    """

df_orders, current_sha = nacti_objednavky()
nastaveni_app, sha_nastaveni = nacti_nastaveni()

# Převod seznamu partnerů na slovník pro rychlé vyhledávání podle PINu
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
        
        col_katalog, col_kosik = st.columns([1.3, 1])
        
        vybrane_polozky = []
        celkova_cena_vo = 0
        celkem_ks = 0
        mo_hodnota_celkem = 0
        
        with col_katalog:
            st.subheader("1. Výběr svíček a produktů")
            st.caption(f"Minimální odběr pro VO cenu je {MIN_OBJEDNAVKA_KC:,.0f} Kč.")
            
            col_p1, col_p2 = st.columns(2)
            
            for idx, prod in enumerate(PRODUKTY_KATALOG):
                target_col = col_p1 if idx % 2 == 0 else col_p2
                cena_mo = prod["cena_mo"]
                cena_vo = round(cena_mo * (1 - aktualni_sleva / 100))
                
                label = f"{prod['nazev']} (VO: {cena_vo} Kč | MO: {cena_mo:.0f} Kč)"
                ks = target_col.number_input(label, min_value=0, max_value=500, value=0, key=f"vo_{idx}")
                
                if ks > 0:
                    vybrane_polozky.append(f"{ks}x {prod['nazev']}")
                    celkova_cena_vo += ks * cena_vo
                    mo_hodnota_celkem += ks * cena_mo
                    celkem_ks += ks

        with col_kosik:
            st.subheader("2. Údaje odběratele & Doručení")
            
            firma = st.text_input("Název firmy / Název obchodu & IČO *", value=nazev_partnera).strip()
            jmeno = st.text_input("Jméno kontaktní osoby *").strip()
            col_t, col_e = st.columns(2)
            telefon = col_t.text_input("Telefon *").strip()
            email = col_e.text_input("E-mail *").strip()
            adresa = st.text_area("Doručovací adresa (Ulice, ČP, Město, PSČ) *", height=80).strip()
            poznamka = st.text_input("Poznámka k doručení / balení:").strip()
            
            st.divider()
            
            uspora = mo_hodnota_celkem - celkova_cena_vo
            
            st.markdown(f"""
                <div class='summary-card'>
                    <h4 style='margin: 0;'>Shrnutí VO objednávky</h4>
                    <p style='margin: 5px 0 0 0; font-size: 14px;'>Celkem produktů: <strong>{celkem_ks} ks</strong></p>
                    <p style='margin: 0; font-size: 13px;'>Vaše úspora oproti MO: {uspora:,.0f} Kč</p>
                    <h3 style='margin: 10px 0 0 0;'>Celkem VO cena: {celkova_cena_vo:,.0f} Kč</h3>
                </div>
            """, unsafe_allow_html=True)
            
            if celkova_cena_vo < MIN_OBJEDNAVKA_KC and celkova_cena_vo > 0:
                st.warning(f"⚠️ Minimální hodnota velkoobchodní objednávky je {MIN_OBJEDNAVKA_KC} Kč. Chybí ještě {MIN_OBJEDNAVKA_KC - celkova_cena_vo} Kč.")
            
            if st.button("Odeslat velkoobchodní objednávku", type="primary", use_container_width=True):
                if not all([firma, jmeno, telefon, email, adresa]):
                    st.warning("⚠️ Prosím vyplňte všechny kontaktní a firemní údaje.")
                elif celkova_cena_vo < MIN_OBJEDNAVKA_KC:
                    st.error(f"❌ Minimální částka objednávky není splněna ({MIN_OBJEDNAVKA_KC} Kč).")
                else:
                    with st.spinner('Odesílám VO objednávku... 🕯️'):
                        nove_id = 1 if df_orders.empty else int(df_orders["ID"].max()) + 1
                        polozky_text = ", ".join(vybrane_polozky)
                        
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
                            "Polozky_Detail": polozky_text,
                            "Celkem_Ks": celkem_ks,
                            "Sleva_Pouzita": f"{aktualni_sleva} %",
                            "Cena_Celkem_VO": celkova_cena_vo,
                            "Stav_Platby": "Čeká na platbu"
                        }])
                        
                        df_aktualni = pd.concat([df_orders, nova_obj], ignore_index=True)
                        if uloz_objednavky(df_aktualni, current_sha):
                            st.balloons()
                            st.success("🎉 Děkujeme! Velkoobchodní objednávka byla úspěšně přijata.")
                            
                            html_uct = vygeneruj_b2b_uctenku(nove_id, firma, jmeno, adresa, telefon, polozky_text, celkem_ks, celkova_cena_vo, aktualni_sleva)
                            st.download_button("📥 Stáhnout B2B Potvrzení (HTML/PDF)", html_uct, file_name=f"DeniCandle_B2B_{nove_id}.html", mime="text/html")
                            
                            spd_str = f"SPD*1.0*ACC:{BANK_ACCOUNT}/{BANK_CODE}*AM:{celkova_cena_vo:.2f}*CC:CZK*X-VS:{nove_id}*MSG:DeniCandle B2B {nove_id}"
                            qr_img = qrcode.make(spd_str)
                            buf = BytesIO()
                            qr_img.save(buf, format="PNG")
                            st.image(buf.getvalue(), caption="QR Platba převodem", width=200)
                        else:
                            st.error("❌ Chyba při ukládání objednávky.")
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
                st.dataframe(df_orders, use_container_width=True, hide_index=True)
            else:
                st.info("Zatím žádné B2B objednávky.")
                
        with tab_vyroba:
            st.subheader("🕯️ Celkový počet svíček k odlítí")
            if not df_orders.empty:
                vsechny_polozky = []
                for detail in df_orders["Polozky_Detail"].dropna():
                    for item in detail.split(", "):
                        if "x " in item:
                            ks, název = item.split("x ", 1)
                            vsechny_polozky.append({"Produkt": název, "Ks": int(ks)})
                
                if vsechny_polozky:
                    df_sum = pd.DataFrame(vsechny_polozky).groupby("Produkt").sum().reset_index()
                    st.table(df_sum)
                else:
                    st.info("Žádné svíčky k výrobě.")
            else:
                st.info("Zatím žádné objednávky v databázi.")
                
        with tab_partneri:
            st.subheader("👥 Správa přístupů a individuálních slev")
            st.write("Vytvořte svým partnerům vlastní PINy a nastavte jim přesnou výši slevy.")
            
            df_partneri = pd.DataFrame(partneri_seznam)
            
            edited_partneri = st.data_editor(
                df_partneri,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "PIN": st.column_config.TextColumn("Přístupový PIN (heslo)", required=True),
                    "Nazev": st.column_config.TextColumn("Název partnera", required=True),
                    "Sleva": st.column_config.NumberColumn("Sleva v %", min_value=0, max_value=99, required=True, format="%d %%")
                }
            )
            
            if st.button("💾 Uložit seznam partnerů a slev", type="primary"):
                nove_nastaveni = {"partneri": edited_partneri.to_dict('records')}
                if uloz_nastaveni(nove_nastaveni, sha_nastaveni):
                    st.success("✅ Úspěšně uloženo! Partneři se nyní mohou přihlásit svými PINy.")
                    st.rerun()
                else:
                    st.error("❌ Při ukládání došlo k chybě.")
