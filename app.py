streamlit
pandas
openpyxl
fpdf

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Gestionale Banca Centro Lazio", page_icon="🏦")
COLORE_AZIENDALE = "#0056b3"

# --- FUNZIONI ---
def carica_dati():
    if os.path.exists("database.xlsx"):
        df = pd.read_excel("database.xlsx")
        # Crea etichetta per ricerca
        df['Etichetta'] = df['Nome'].astype(str) + " " + df['Cognome'].astype(str)
        return df
    return None

def crea_pdf_in_memoria(dati):
    # Inizializza PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # 1. INSERIMENTO LOGO NEL PDF (Se esiste)
    if os.path.exists("logo.jpg"):
        # x=10, y=8, w=30 (dimensioni e posizione logo nel foglio)
        pdf.image("logo.jpg", x=10, y=8, w=30)
        pdf.ln(20) # Vai a capo dopo il logo

    # 2. TITOLO
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, txt="Scheda Personale - Banca Centro Lazio", ln=True, align='C')
    pdf.ln(10) # Spazio vuoto

    # 3. DATI DEL DIPENDENTE
    pdf.set_font("Arial", size=12)
    
    # Funzione per scrivere riga: Etichetta in grassetto, Valore normale
    def scrivi_riga_pdf(etichetta, valore):
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(50, 10, etichetta, border=0)
        pdf.set_font("Arial", '', 12)
        # Usiamo latin-1 per gestire gli accenti italiani
        valore_str = str(valore).encode('latin-1', 'replace').decode('latin-1')
        pdf.cell(0, 10, valore_str, border=0, ln=True)

    scrivi_riga_pdf("Nome e Cognome:", f"{dati['Nome']} {dati['Cognome']}")
    scrivi_riga_pdf("Ruolo:", dati['Ruolo'])
    scrivi_riga_pdf("Email:", dati['Email'])
    tel = dati['Telefono'] if 'Telefono' in dati else "N/D"
    scrivi_riga_pdf("Telefono:", tel)

    # 4. TESTO LEGALE
    pdf.ln(10)
    pdf.multi_cell(0, 10, "Si certifica che i dati sopra riportati sono corretti e verificati dal sistema.")

    # 5. FIRME
    pdf.ln(30) # Molto spazio prima della firma
    
    data_oggi = datetime.date.today().strftime('%d/%m/%Y')
    
    # Riga Data
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(20, 10, "Data: ", border=0)
    pdf.set_font("Arial", '', 12)
    pdf.cell(80, 10, data_oggi, border=0)
    
    # Riga Firma
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(40, 10, "Firma Responsabile:", border=0)
    pdf.cell(0, 10, "_"*30, border=0, ln=True) # Linea tratteggiata

    # Restituisce il contenuto PDF come stringa binaria
    return pdf.output(dest='S').encode('latin-1')

# --- INTERFACCIA WEB ---

# Intestazione con Logo (VISIBILE SUL SITO)
col1, col2 = st.columns([1, 4])
with col1:
    if os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=100)
    else:
        st.write("🏦")
with col2:
    st.title("Banca Centro Lazio")
    st.caption("Portale Gestione Risorse Umane")

st.divider()

# Logica Applicazione
df = carica_dati()

if df is None:
    st.error("⚠️ File 'database.xlsx' mancante!")
else:
    # Menu a tendina con ricerca
    lista_nomi = ["Seleziona un candidato..."] + df['Etichetta'].tolist()
    
    scelta_utente = st.selectbox(
        "🔍 Cerca dipendente (inizia a scrivere):",
        options=lista_nomi
    )

    if scelta_utente != "Seleziona un candidato...":
        riga = df[df['Etichetta'] == scelta_utente].iloc[0]
        
        st.success("Dipendente trovato")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            st.subheader(f"{riga['Nome']} {riga['Cognome']}")
            st.write(f"**Ruolo:** {riga['Ruolo']}")
            st.write(f"**Email:** {riga['Email']}")
        
        with c2:
            st.write("")
            # Genera PDF
            file_pdf = crea_pdf_in_memoria(riga)
            
            # Tasto Download PDF
            st.download_button(
                label="📥 Scarica PDF",
                data=file_pdf,
                file_name=f"Scheda_{riga['Nome']}_{riga['Cognome']}.pdf",
                mime="application/pdf",
                type="primary"
            )
