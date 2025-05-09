import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name="Consulente")
    # Sezione InserimentoGruppi
    grp_df = (
        cfg[cfg["Section"] == "InserimentoGruppi"]
        [["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    )
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))

    # Sezione Defaults
    def_df = (
        cfg[cfg["Section"] == "Defaults"]
        [["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "key", "Label/Gruppi/Value": "value"})
    )
    defaults = dict(zip(def_df["key"], def_df["value"]))

    return gruppi, defaults

# ------------------------------------------------------------
# App 1.3: Risorsa Esterna - Consulente
# ------------------------------------------------------------
st.set_page_config(page_title="1.3 Risorsa Esterna: Consulente")
st.title("1.3 Risorsa Esterna: Consulente")

config_file = st.file_uploader(
    "Carica il file di configurazione (config_corrected.xlsx)",
    type=["xlsx"],
    help="Deve contenere il foglio “Consulente” con colonne Section, Key/App, Label/Gruppi/Value"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per procedere.")
    st.stop()

gruppi, defaults = load_config_from_bytes(config_file.read())

# ------------------------------------------------------------
# Determinazione del valore di OU dal default (ou_default)
# ------------------------------------------------------------
ou_value = defaults.get("ou_default", "")

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------
def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

def genera_samaccountname(nome: str, cognome: str,
                          secondo_nome: str = "", secondo_cognome: str = "",
                          esterno: bool = False) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit  = 16 if esterno else 20
    cand   = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    cand = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}")[:limit] + suffix

def build_full_name(cognome: str, secondo_cognome: str,
                    nome: str, secondo_nome: str,
                    esterno: bool = False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full  = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]

# ------------------------------------------------------------
# Form di input nell’ordine richiesto
# ------------------------------------------------------------
cognome         = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
nome            = st.text_input("Nome").strip().capitalize()
secondo_nome    = st.text_input("Secondo Nome").strip().capitalize()
cf              = st.text_input("Codice Fiscale", "").strip()
telefono        = st.text_input("Mobile", "").replace(" ", "")
description     = st.text_input("PC", "<PC>").strip()
exp_date        = st.text_input(
    "Data di Fine (gg-mm-aaaa)",
    defaults.get("expire_default", "30-06-2025")
).strip()

email_flag = st.radio("Email Consip necessaria?", ["Sì", "No"]) == "Sì"
if not email_flag:
    custom_email = st.text_input("Email Personalizzata", "").strip()
else:
    custom_email = None

# ------------------------------------------------------------
# Valori fissi prelevati dalla configurazione
# ------------------------------------------------------------
employee_id        = defaults.get("employee_id_default", "")
department         = defaults.get("department_consulente", "")
inserimento_gruppo = gruppi.get("esterna_consulente", "")
company            = defaults.get("company_default", "")

# ------------------------------------------------------------
# Generazione CSV
# ------------------------------------------------------------
if st.button("Genera CSV Consulente"):
    sAM     = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn      = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt = formatta_data(exp_date)
    upn     = f"{sAM}@consip.it"
    mail    = upn if email_flag else (custom_email or upn)
    mobile  = f"+39 {telefono}" if telefono else ""
    given   = f"{nome} {secondo_nome}".strip()
    surn    = f"{cognome} {secondo_cognome}".strip()

    row = [
        sAM, "SI", ou_value, cn, cn, cn, given, surn,
        cf, employee_id, department, description or "<PC>", "No", exp_fmt,
        upn, mail, mobile, "", inserimento_gruppo, "", "",
        "", company  # telephoneNumber always empty
    ]

    buf = io.StringIO()
    # Disattiviamo il quoting automatico e definiamo "\\" come escapechar
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")

    # Aggiungiamo manualmente i doppi apici ai campi indicati
    for i in (2, 3, 4, 5):       # OU, Name, DisplayName, cn
        row[i] = f"\"{row[i]}\""
    if secondo_nome:             # GivenName solo se presente
        row[6] = f"\"{row[6]}\""
    if secondo_cognome:          # Surname solo se presente
        row[7] = f"\"{row[7]}\""
    row[13] = f"\"{row[13]}\""   # ExpireDate
    row[16] = f"\"{row[16]}\""   # mobile

    # Scriviamo header + riga
    writer.writerow(HEADER)
    writer.writerow(row)
    buf.seek(0)

    st.dataframe(pd.DataFrame([row], columns=HEADER))
    st.download_button(
        label="📥 Scarica CSV Consulente",
        data=buf.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_consulente.csv",
        mime="text/csv"
    )
    st.success(f"✅ File CSV generato per '{sAM}'")
