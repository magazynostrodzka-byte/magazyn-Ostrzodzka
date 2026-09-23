import os
import io
import smtplib
from email.message import EmailMessage
from datetime import datetime
import pandas as pd
import streamlit as st
from streamlit_drawable_canvas import st_canvas
from PIL import Image

# Biblioteki do generowania PDF (ReportLab)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==========================================
# KONFIGURACJA BEZPIECZEŃSTWA I POCZTY
# ==========================================
HASLO_MAGAZYNU = "Magazyn2026!" # Hasło wymagane przy logowaniu do aplikacji

LISTA_BUDOW = ["M422- Ostródzka IV/Wawel", "M9-Klaudyn ul. Ekologiczna 4"]
LISTA_ODBIORCOW = ["Magazyn Własny", "Podwykonawca - Zbrojenia", "Podwykonawca - Cieśle"]
LISTA_DOSTAWCOW = ["Dostawa Zewnętrzna", "Przesunięcie z innej budowy"]

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
SENDER_EMAIL = "magazyn.ostrodzka@gmail.com"
SENDER_PASSWORD = "xbjv onpy kzqk ldvi" # Do uzupełnienia hasłem z Google

# ==========================================
# KONFIGURACJA STRONY I PLIKÓW
# ==========================================
st.set_page_config(page_title="Magazyn Budowlany", page_icon="🏗️", layout="wide")

FILE_STAN = "stan_magazynowy.xlsx"
FILE_PRZYJAZDY = "lista_przyjazdow.xlsx"
FILE_WYJAZDY = "lista_wyjazdow.xlsx"
FILE_LOGI = "pelne_logi_operacji.xlsx"
FOLDER_ARCHIWUM = "ARCHIWUM_DOKUMENTOW"
os.makedirs(FOLDER_ARCHIWUM, exist_ok=True)

def wyslij_email_z_pdf(odbiorca_email, nazwa_pliku, pdf_bytes):
    try:
        if SENDER_PASSWORD == "wpisz_tutaj_haslo_aplikacji":
            st.warning("E-mail nie został wysłany: Brak skonfigurowanego hasła aplikacji.")
            return False
            
        msg = EmailMessage()
        msg['Subject'] = f"Kopia dokumentu: {nazwa_pliku}"
        msg['From'] = SENDER_EMAIL
        msg['To'] = odbiorca_email
        msg.set_content(f"Cześć,\n\nW załączniku przesyłamy dokument wygenerowany w Systemie Magazynowym: {nazwa_pliku}.")
        msg.add_attachment(pdf_bytes.getvalue(), maintype='application', subtype='pdf', filename=nazwa_pliku)
        
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Nie udało się wysłać e-maila. Błąd: {e}")
        return False

# ==========================================
# INICJALIZACJA BAZY DANYCH
# ==========================================
def init_excel_files():
    kolumny_przyjazdow = ["Data_Godzina", "Wystawiajacy", "Z_Skad", "Do_Dokad", "Uwagi", "Kod", "Nazwa", "Nr_Seryjny", "Ilosc"]
    kolumny_wyjazdow = ["Data_Godzina", "Nr_Dokumentu", "Wystawiajacy", "Z_Skad", "Do_Dokad", "Uwagi", "Kod", "Nazwa", "Ilosc"]
    kolumny_stan = ["Kod", "Nazwa", "Stan"]
    
    if not os.path.exists(FILE_PRZYJAZDY):
        pd.DataFrame(columns=kolumny_przyjazdow).to_excel(FILE_PRZYJAZDY, index=False)
    if not os.path.exists(FILE_WYJAZDY):
        pd.DataFrame(columns=kolumny_wyjazdow).to_excel(FILE_WYJAZDY, index=False)
    if not os.path.exists(FILE_LOGI):
        pd.DataFrame(columns=kolumny_wyjazdow).to_excel(FILE_LOGI, index=False)
    if not os.path.exists(FILE_STAN):
        pd.DataFrame(columns=kolumny_stan).to_excel(FILE_STAN, index=False)

init_excel_files()

def clean_code(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str.upper()

def is_x_value(val):
    if pd.isna(val): return False
    return str(val).strip().upper() == "X"

def get_przyjazdy(): 
    try: return pd.read_excel(FILE_PRZYJAZDY)
    except: return pd.DataFrame(columns=["Data_Godzina", "Wystawiajacy", "Z_Skad", "Do_Dokad", "Uwagi", "Kod", "Nazwa", "Nr_Seryjny", "Ilosc"])

def save_przyjazdy(df): df.to_excel(FILE_PRZYJAZDY, index=False)

def get_wyjazdy(): 
    try: return pd.read_excel(FILE_WYJAZDY)
    except: return pd.DataFrame(columns=["Data_Godzina", "Nr_Dokumentu", "Wystawiajacy", "Z_Skad", "Do_Dokad", "Uwagi", "Kod", "Nazwa", "Ilosc"])

def save_wyjazdy(df): df.to_excel(FILE_WYJAZDY, index=False)

def get_logi(): 
    try: return pd.read_excel(FILE_LOGI)
    except: return pd.DataFrame(columns=["Data_Godzina", "Nr_Dokumentu", "Wystawiajacy", "Z_Skad", "Do_Dokad", "Uwagi", "Kod", "Nazwa", "Ilosc"])

def save_logi(df): df.to_excel(FILE_LOGI, index=False)

# ==========================================
# PRECYZYJNE PRZELICZANIE STANU (SYMBIOZA)
# ==========================================
def oblicz_i_zapisz_aktualny_stan():
    df_pz = get_przyjazdy()
    df_wz = get_wyjazdy()
    magazyn = {}

    # 1. Sumowanie wszystkich przyjazdów
    if not df_pz.empty:
        for _, row in df_pz.iterrows():
            kod = clean_code(row.get("Kod"))
            if not kod: continue
            nazwa = str(row.get("Nazwa", "Brak nazwy")).strip()
            
            try: ilosc = int(float(str(row.get("Ilosc", 0)).replace(',', '.')))
            except: ilosc = 0

            if kod not in magazyn:
                magazyn[kod] = {"Kod": kod, "Nazwa": nazwa, "Stan": 0}
            magazyn[kod]["Stan"] += ilosc
            if nazwa and nazwa != "nan":
                magazyn[kod]["Nazwa"] = nazwa

    # 2. Odejmowanie wszystkich wyjazdów (z pominięciem "X")
    if not df_wz.empty:
        for _, row in df_wz.iterrows():
            kod = clean_code(row.get("Kod"))
            if not kod: continue
            nazwa = str(row.get("Nazwa", "Brak nazwy")).strip()
            ilosc_val = row.get("Ilosc")
            
            ilosc = 0
            if not is_x_value(ilosc_val):
                try: ilosc = int(float(str(ilosc_val).replace(',', '.')))
                except: ilosc = 0

                if kod not in magazyn:
                    magazyn[kod] = {"Kod": kod, "Nazwa": nazwa, "Stan": 0}
                magazyn[kod]["Stan"] -= ilosc
                if nazwa and nazwa != "nan":
                    magazyn[kod]["Nazwa"] = nazwa

    dane = list(magazyn.values())
    df_stan = pd.DataFrame(dane) if dane else pd.DataFrame(columns=["Kod", "Nazwa", "Stan"])
    df_stan = df_stan.sort_values(by="Kod").reset_index(drop=True)
    df_stan.to_excel(FILE_STAN, index=False)
    
    return df_stan[df_stan["Stan"] > 0].reset_index(drop=True)

def get_stan_magazynowy():
    return oblicz_i_zapisz_aktualny_stan()

def get_wszystkie_materialy_bazy():
    materialy = {}
    df_pz = get_przyjazdy()
    df_wz = get_wyjazdy()

    if not df_pz.empty:
        for _, r in df_pz.iterrows():
            k, n = clean_code(r.get("Kod")), r.get("Nazwa")
            if k and not pd.isna(n):
                materialy[k] = str(n).strip()

    if not df_wz.empty:
        for _, r in df_wz.iterrows():
            k, n = clean_code(r.get("Kod")), r.get("Nazwa")
            if k and not pd.isna(n):
                materialy[k] = str(n).strip()

    return materialy

# ==========================================
# GENEROWANIE DOKUMENTU PDF
# ==========================================
def generate_pdf(nr_dok, z_skad, do_dokad, data_dok, pozycje, signature_img_bytes=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=40, bottomMargin=70)
    story, styles = [], getSampleStyleSheet()
    
    try:
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
        pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
        pdfmetrics.registerFont(TTFont('Arial-Italic', 'ariali.ttf'))
        pdfmetrics.registerFont(TTFont('Arial-BoldItalic', 'arialbi.ttf'))
        
        FONT_NORMAL = 'Arial'
        FONT_BOLD = 'Arial-Bold'
        FONT_TITLE = 'Arial-BoldItalic'
    except:
        FONT_NORMAL = 'Helvetica'
        FONT_BOLD = 'Helvetica-Bold'
        FONT_TITLE = 'Helvetica-BoldOblique'
    
    style_normal = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName=FONT_NORMAL, fontSize=10, leading=14)
    style_bold = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=10, leading=14)
    style_title = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=FONT_TITLE, fontSize=16, alignment=1, spaceAfter=5)
    style_subtitle = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontName=FONT_NORMAL, fontSize=9, alignment=1, spaceAfter=20)
    
    logo_path = "logo.png" if os.path.exists("logo.png") else ("logo.jpg" if os.path.exists("logo.jpg") else None)
    if logo_path:
        img_logo = RLImage(logo_path, width=120, height=45)
        img_logo.hAlign = 'LEFT'
        story.append(img_logo)
        story.append(Spacer(1, 10))

    story.append(Paragraph("DOKUMENT MAGAZYNOWY", style_title))
    story.append(Paragraph("MALBUD1 Sp. z o.o. Sp.k ul. Obozowa 57 01-161 Warszawa tel.: +48 22 664 69 57 biuro@malbud1.pl", style_subtitle))
    story.append(Spacer(1, 10))

    meta_table_data = [
        [Paragraph("<b>Numer dokumentu:</b>", style_normal), Paragraph(nr_dok, style_normal)],
        [Paragraph("<b>Typ dokumentu:</b>", style_normal), Paragraph("Przewozowy", style_normal)],
        [Paragraph("<b>Przewóz z:</b>", style_normal), Paragraph(z_skad, style_normal)],
        [Paragraph("<b>Przewóz do:</b>", style_normal), Paragraph(do_dokad, style_normal)],
        [Paragraph("<b>Data:</b>", style_normal), Paragraph(data_dok, style_normal)]
    ]
    
    meta_table = Table(meta_table_data, colWidths=[120, 375])
    meta_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 15))

    table_data = []
    for item in pozycje:
        nazwa_wyswietlana = f"{item['Nazwa']} (SN: {item['Nr_Seryjny']})" if item.get('Nr_Seryjny') else item['Nazwa']
        table_data.append([Paragraph(str(item["Kod"]), style_normal), Paragraph(str(nazwa_wyswietlana), style_normal), Paragraph(str(item['Ilosc']), style_normal)])
    
    if not table_data: table_data.append(["", "Brak pozycji", ""])

    items_table = Table(table_data, colWidths=[110, 335, 50])
    items_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(items_table)
    story.append(Spacer(1, 30))

    podpis_img = RLImage(io.BytesIO(signature_img_bytes), width=120, height=40) if signature_img_bytes else Paragraph("<br/><br/><br/>", style_normal)
    
    sig_data = [
        [podpis_img, Paragraph("<br/><br/><br/>", style_normal)],
        [Paragraph("Podpis osoby wysyłającej", style_normal), Paragraph("Podpis osoby odbierającej", style_normal)]
    ]
    
    sig_table = Table(sig_data, colWidths=[247, 248])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('ALIGN', (1,0), (1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(sig_table)

    def add_footer(canvas, doc):
        canvas.saveState()
        stopka_path = "stopka.png" if os.path.exists("stopka.png") else ("stopka.jpg" if os.path.exists("stopka.jpg") else None)
        if stopka_path:
            canvas.drawImage(stopka_path, 50, 15, width=495, height=35, preserveAspectRatio=True, mask='auto')
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    buffer.seek(0)
    return buffer

# ==========================================
# EKRAN LOGOWANIA / IDENTYFIKACJI + HASŁO
# ==========================================
if "user_identified" not in st.session_state: st.session_state.user_identified = False

if not st.session_state.user_identified:
    st.title("🔒 Logowanie do Magazynu Budowlanego")
    with st.form("login_form"):
        imie = st.text_input("Imię")
        nazwisko = st.text_input("Nazwisko")
        email = st.text_input("Adres e-mail")
        haslo_input = st.text_input("Hasło dostępu do magazynu", type="password")
        
        if st.form_submit_button("Wejdź do magazynu"):
            if not (imie and nazwisko and email and haslo_input):
                st.error("Wypełnij wszystkie pola formularza.")
            elif haslo_input != HASLO_MAGAZYNU:
                st.error("❌ Niepoprawne hasło dostępu do magazynu!")
            else:
                st.session_state.user_identified = True
                st.session_state.imie, st.session_state.nazwisko, st.session_state.email = imie.strip(), nazwisko.strip(), email.strip()
                st.rerun()
    st.stop()

# ==========================================
# GŁÓWNY INTERFEJS
# ==========================================
pelne_nazwisko = f"{st.session_state.imie} {st.session_state.nazwisko}"
st.title(f"🏗️ Magazyn (Zalogowano: {pelne_nazwisko})")

tab_przyjazd, tab_wyjazd, tab_korekta, tab_stan, tab_logi = st.tabs(["📥 Przyjazd", "🚚 Wyjazd (z opcją X)", "🛠️ Korekta / Rozliczenie X", "📦 Stan Magazynowy", "📊 Logi / Historia"])

# ------------------------------------------
# TAB 1: PRZYJAZD
# ------------------------------------------
with tab_przyjazd:
    col_pz1, col_pz2 = st.columns(2)
    with col_pz1:
        skad_pz = st.selectbox("Z (Skąd przyjechało):", LISTA_DOSTAWCOW + LISTA_BUDOW, key="skad_pz")
        dokad_pz = st.selectbox("DO (Dokąd trafiło):", LISTA_BUDOW, key="dokad_pz")
    with col_pz2:
        nr_rej_pz = st.text_input("Nr rejestracyjny / Auto:", key="nr_rej_pz")
        uwagi_pz = st.text_input("Uwagi / Nr dok. dostawcy:", key="uwagi_pz")

    st.markdown("---")
    st.markdown("### Dodaj pozycję na przyjazd")
    
    tryb_dodawania = st.radio("Rodzaj materiału:", ["Wybierz z bazy materiałów", "Dodaj zupełnie nowy kod"], horizontal=True)
    baza_materialow = get_wszystkie_materialy_bazy()
    
    if "basket_pz" not in st.session_state: st.session_state.basket_pz = []
    
    kod_dodawany = ""
    nazwa_dodawana = ""
    nr_seryjny_pz = ""
    ilosc_pz = 1

    c1, c2, c3, c4 = st.columns([1, 2, 1.5, 1])
    if tryb_dodawania == "Wybierz z bazy materiałów":
        opcje = [""] + [f"{kod} | {nazwa}" for kod, nazwa in baza_materialow.items()]
        wybrany_mat = st.selectbox("Wyszukaj materiał z bazy:", options=opcje, key="sel_pz")
        if wybrany_mat:
            kod_dodawany = wybrany_mat.split(" | ")[0]
            nazwa_dodawana = baza_materialow[kod_dodawany]
            with c3: nr_seryjny_pz = st.text_input("Nr seryjny (opcjonalnie):", key="sn_pz_lista")
            with c4: ilosc_pz = st.number_input("Ilość (szt.):", min_value=1, key="il_pz_lista")
    else:
        with c1: kod_dodawany = st.text_input("Kod towaru:", key="nowy_kod_pz")
        with c2: nazwa_dodawana = st.text_input("Nazwa towaru:", key="nowa_nazwa_pz")
        with c3: nr_seryjny_pz = st.text_input("Nr seryjny (opcjonalnie):", key="sn_pz_nowy")
        with c4: ilosc_pz = st.number_input("Ilość (szt.):", min_value=1, key="il_pz_nowy")

    if kod_dodawany and nazwa_dodawana:
        if st.button("➕ Dodaj do koszyka przyjazdu"):
            st.session_state.basket_pz.append({"Kod": clean_code(kod_dodawany), "Nazwa": nazwa_dodawana, "Nr_Seryjny": nr_seryjny_pz, "Ilosc": ilosc_pz})
            st.rerun()

    if st.session_state.basket_pz:
        st.dataframe(pd.DataFrame(st.session_state.basket_pz), use_container_width=True)
        
        if st.button("📥 ZATWIERDŹ PRZYJAZD (Zapisz w bazie)", type="primary"):
            df_przyjazdy = get_przyjazdy()
            pelne_z = f"{skad_pz} {nr_rej_pz}".strip()
            nowe_wpisy = []

            for item in st.session_state.basket_pz:
                nowe_wpisy.append({
                    "Data_Godzina": datetime.now().strftime("%Y-%m-%d %H:%M"), 
                    "Wystawiajacy": pelne_nazwisko, 
                    "Z_Skad": pelne_z, 
                    "Do_Dokad": dokad_pz, 
                    "Uwagi": uwagi_pz, 
                    "Kod": clean_code(item["Kod"]), 
                    "Nazwa": item["Nazwa"], 
                    "Nr_Seryjny": item["Nr_Seryjny"],
                    "Ilosc": item["Ilosc"]
                })
            
            save_przyjazdy(pd.concat([df_przyjazdy, pd.DataFrame(nowe_wpisy)], ignore_index=True))
            oblicz_i_zapisz_aktualny_stan()

            st.success("✅ Przyjazd został pomyślnie zarejestrowany w systemie!")
            st.session_state.basket_pz = []
            st.rerun()

# ------------------------------------------
# TAB 2: WYJAZD
# ------------------------------------------
with tab_wyjazd:
    col_wz1, col_wz2 = st.columns(2)
    with col_wz1:
        skad_wz = st.selectbox("Z (Skąd wyjeżdża):", LISTA_BUDOW, key="skad_wz")
        dokad_wz = st.selectbox("DO (Odbiorca / Cel):", LISTA_BUDOW + LISTA_ODBIORCOW, key="dokad_wz")
    with col_wz2:
        nr_rej_wz = st.text_input("Nr rejestracyjny / Auto:", key="nr_rej_wz")
        uwagi_wz = st.text_input("Uwagi:", key="uwagi_wz")

    df_mag_wz = get_stan_magazynowy()
    if "basket_wz" not in st.session_state: st.session_state.basket_wz = []
    
    opcje_wz = [""] + [f"{clean_code(r['Kod'])} | {r['Nazwa']} (Stan: {r['Stan']} szt.)" for _, r in df_mag_wz.iterrows()] if not df_mag_wz.empty else [""]

    wybrany_mat_wz = st.selectbox("Wybierz materiał do wyjazdu:", options=opcje_wz)
    
    if wybrany_mat_wz:
        kod_wybrany = clean_code(wybrany_mat_wz.split(" | ")[0])
        mat_row = df_mag_wz[df_mag_wz["Kod"].apply(clean_code) == kod_wybrany].iloc[0]
        c1, c2 = st.columns([2, 1])
        with c1: ilosc_wz_input = st.text_input("Ilość do wyjazdu (liczba lub 'X'):", value="1", key="il_wz_str")
        with c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕ Dodaj do koszyka wyjazdu"):
                ilosc_zapis = ilosc_wz_input.strip().upper()
                if ilosc_zapis != "X":
                    try:
                        ilosc_zapis = int(float(ilosc_zapis.replace(',', '.')))
                        if mat_row["Stan"] < ilosc_zapis:
                            st.error("Brak wystarczającej ilości w magazynie!")
                            st.stop()
                    except:
                        ilosc_zapis = "X"

                st.session_state.basket_wz.append({"Kod": kod_wybrany, "Nazwa": mat_row["Nazwa"], "Ilosc": ilosc_zapis})
                st.rerun()

    if st.session_state.basket_wz:
        st.dataframe(pd.DataFrame(st.session_state.basket_wz), use_container_width=True)
        st.write("Twój podpis (Osoba wysyłająca):")
        canvas_result_wz = st_canvas(stroke_width=2, stroke_color="#000", background_color="#FFF", height=120, width=400, drawing_mode="freedraw", key="canvas_wz", return_image_data=True)

        if st.button("🚚 ZATWIERDŹ WYJAZD I WYŚLIJ DOKUMENT", type="primary"):
            sig_bytes = None
            if canvas_result_wz.image_data is not None:
                img = Image.fromarray(canvas_result_wz.image_data.astype('uint8'), 'RGBA')
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                sig_bytes = buf.getvalue()

            df_wyjazdy = get_wyjazdy()
            df_logi = get_logi()
            nr_dok = f"P/{datetime.now().strftime('%Y/%m')}/{(len(df_wyjazdy) + 1):03d}"
            data_dok = datetime.now().strftime('%d-%m-%Y')
            pelny_cel = f"{dokad_wz} {nr_rej_wz}".strip()
            nowe_wpisy = []
            
            for item in st.session_state.basket_wz:
                nowe_wpisy.append({
                    "Data_Godzina": datetime.now().strftime("%Y-%m-%d %H:%M"), 
                    "Nr_Dokumentu": nr_dok, 
                    "Wystawiajacy": pelne_nazwisko, 
                    "Z_Skad": skad_wz, 
                    "Do_Dokad": pelny_cel, 
                    "Uwagi": uwagi_wz, 
                    "Kod": clean_code(item["Kod"]), 
                    "Nazwa": item["Nazwa"], 
                    "Ilosc": item["Ilosc"]
                })
            
            df_nowe = pd.DataFrame(nowe_wpisy)
            save_wyjazdy(pd.concat([df_wyjazdy, df_nowe], ignore_index=True))
            save_logi(pd.concat([df_logi, df_nowe], ignore_index=True))
            oblicz_i_zapisz_aktualny_stan()

            pdf_bytes = generate_pdf(nr_dok, skad_wz, pelny_cel, data_dok, st.session_state.basket_wz, sig_bytes)
            nazwa_pdf = f"WZ_{datetime.now().strftime('%Y-%m-%d')}_{st.session_state.nazwisko.replace(' ', '_')}_{len(df_wyjazdy)+1}.pdf"
            with open(os.path.join(FOLDER_ARCHIWUM, nazwa_pdf), "wb") as f: f.write(pdf_bytes.getbuffer())

            with st.spinner("Wysyłanie dokumentu na e-mail..."): wyslij_email_z_pdf(st.session_state.email, nazwa_pdf, pdf_bytes)
            st.success(f"✅ Towar wydano! Dokument {nr_dok} zapisany pomyślnie.")
            st.download_button(label="📥 Pobierz Dokument", data=pdf_bytes, file_name=nazwa_pdf, mime="application/pdf")
            st.session_state.basket_wz = []

# ------------------------------------------
# TAB 3: KOREKTA / ROZLICZENIE X
# ------------------------------------------
with tab_korekta:
    st.subheader("🛠️ Rozliczenie niepoliczonych wyjazdów (zmiana z 'X' na dokładną ilość)")
    
    df_wz_all = get_wyjazdy()
    
    if not df_wz_all.empty and "Ilosc" in df_wz_all.columns:
        maska_x = df_wz_all["Ilosc"].apply(is_x_value)
        df_x = df_wz_all[maska_x]
        
        if not df_x.empty:
            st.markdown("#### ⚠️ Oczekujące pozycje wyjazdowe z 'X':")
            st.dataframe(df_x, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### Wybierz pozycję do poprawienia:")
            
            opcje_x = [f"Wiersz ID: {idx} | Data: {row.get('Data_Godzina')} | Kod: {row.get('Kod')} | Nazwa: {row.get('Nazwa')}" for idx, row in df_x.iterrows()]
            wybrana_pozycja = st.selectbox("Pozycja do rozliczenia:", options=[""] + opcje_x)
            
            if wybrana_pozycja:
                row_idx = int(wybrana_pozycja.split(" | ")[0].replace("Wiersz ID: ", ""))
                wiersz_do_poprawy = df_wz_all.loc[row_idx]
                
                st.info(f"Wybrana pozycja z dnia **{wiersz_do_poprawy.get('Data_Godzina')}**: {wiersz_do_poprawy.get('Kod')} - {wiersz_do_poprawy.get('Nazwa')}")
                
                nowa_ilosc_val = st.number_input("Wpisz FAKTYCZNĄ przeliczoną ilość (szt.):", min_value=1, value=1, key="nowa_ilosc_korekta")
                
                if st.button("⚖️ ZAPISZ ILOŚĆ I POPRAW WIERSZ", type="primary"):
                    df_wz_all.loc[row_idx, "Ilosc"] = int(nowa_ilosc_val)
                    save_wyjazdy(df_wz_all)
                    
                    df_logi = get_logi()
                    nowy_log = pd.DataFrame([{
                        "Data_Godzina": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "Nr_Dokumentu": wiersz_do_poprawy.get("Nr_Dokumentu"),
                        "Wystawiajacy": pelne_nazwisko,
                        "Z_Skad": wiersz_do_poprawy.get("Z_Skad"),
                        "Do_Dokad": wiersz_do_poprawy.get("Do_Dokad"),
                        "Uwagi": f"Korekta ilości X na {nowa_ilosc_val}",
                        "Kod": wiersz_do_poprawy.get("Kod"),
                        "Nazwa": wiersz_do_poprawy.get("Nazwa"),
                        "Ilosc": nowa_ilosc_val
                    }])
                    save_logi(pd.concat([df_logi, nowy_log], ignore_index=True))
                    
                    oblicz_i_zapisz_aktualny_stan()
                    st.success(f"✅ Pozycja zaktualizowana! Zmieniono 'X' na {nowa_ilosc_val} szt.")
                    st.rerun()
        else:
            st.success("🎉 Brak oczekujących wyjazdów z oznaczeniem 'X'! Wszystko jest rozliczone.")
    else:
        st.info("Brak wpisów wyjazdowych w bazie.")

# ------------------------------------------
# TAB 4: STAN MAGAZYNOWY
# ------------------------------------------
with tab_stan:
    st.subheader("Bieżący stan magazynu")
    st.dataframe(get_stan_magazynowy(), use_container_width=True)

# ------------------------------------------
# TAB 5: LOGI / HISTORIA
# ------------------------------------------
with tab_logi:
    st.subheader("📊 Rejestr wykonanych operacji (Logi)")
    typ_logu = st.selectbox("Wybierz historię do wyświetlenia:", ["Wszystkie Wyjazdy / Logi WZ", "Wszystkie Przyjazdy PZ"])
    if typ_logu == "Wszystkie Wyjazdy / Logi WZ":
        st.dataframe(get_logi().sort_values(by="Data_Godzina", ascending=False), use_container_width=True)
    else:
        st.dataframe(get_przyjazdy().sort_values(by="Data_Godzina", ascending=False), use_container_width=True)