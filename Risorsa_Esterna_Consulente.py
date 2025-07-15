import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io
import unicodedata

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name="Consulente")
    # Sezione InserimentoGruppi (per CSV)
    grp_df = (
        cfg[cfg["Section"] == "InserimentoGruppi"][["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    )
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"].astype(str)))

    # Sezione Defaults
    def_df = (
        cfg[cfg["Section"] == "Defaults"][["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "key", "Label/Gruppi/Value": "value"})
    )
    defaults = dict(zip(def_df["key"], def_df["value"].astype(str)))

    return gruppi, defaults

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------
def normalize_name(s: str) -> str:
    """Rimuove spazi, apostrofi e accenti, restituisce in minuscolo."""
    nfkd = unicodedata.normalize('NFKD', s)
    ascii_str = nfkd.encode('ASCII', 'ignore').decode()
    return ascii_str.replace(' ', '').replace("'", '').lower()


def formatta_data(data: str) -> str:
    """Formatta gg-mm-aaaa in MM/DD/YYYY 00:00."""
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

# ------------------------------------------------------------
# Generazione SAMAccountName con normalizzazione
# ------------------------------------------------------------
def genera_samaccountname(
    nome: str,
    cognome: str,
    secondo_nome: str = "",
    secondo_cognome: str = "",
    esterno: bool = False
) -> str:
    n  = normalize_name(nome)
    sn = normalize_name(secondo_nome)
    c  = normalize_name(cognome)
    sc = normalize_name(secondo_cognome)
    suffix = ".ext" if esterno else ""
    limit  = 16 if esterno else 20
    cand1 = f"{n}{sn}.{c}{sc}"
    if len(cand1) <= limit:
        return cand1 + suffix
    cand2 = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand2) <= limit:
        return cand2 + suffix
    base = f"{n[:1]}{sn[:1]}.{c}"
    return base[:limit] + suffix

# ------------------------------------------------------------
# Costruzione display name
# ------------------------------------------------------------
def build_full_name(
    cognome: str,
    secondo_cognome: str,
    nome: str,
    secondo_nome: str,
    esterno: bool = False
) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    return " ".join(parts) + (" (esterno)" if esterno else "")

# ------------------------------------------------------------
# CSV Header
# ------------------------------------------------------------
HEADER = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

# ------------------------------------------------------------
# App 1.3: Risorsa Esterna - Consulente
# ------------------------------------------------------------
st.set_page_config(page_title="1.3 Risorsa Esterna: Consulente")
st.title("1.3 Risorsa Esterna: Consulente")

config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)", type=["xlsx"],
    help="Deve contenere il foglio Consulente con Section, Key/App, Label/Gruppi/Value"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per procedere.")
    st.stop()

gruppi, defaults = load_config_from_bytes(config_file.read())

# ------------------------------------------------------------
# Valori fixed da defaults
# ------------------------------------------------------------
ou_value             = defaults.get("ou_default", "")
department_default   = defaults.get("department_default", "")
description_default  = defaults.get("description_default", "<PC>")
company               = defaults.get("company_default", "")
inserimento_base      = gruppi.get("esterna_consulente", "")
inserimento_noemail   = gruppi.get("esterna_consulente_No_email", "")
o365_std, o365_team, o365_cop = (
    defaults.get("grp_o365_standard", ""),
    defaults.get("grp_o365_teams", ""),
    defaults.get("grp_o365_copilot", "")
)

# ------------------------------------------------------------
# Form di input
# ------------------------------------------------------------
cognome         = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
nome            = st.text_input("Nome").strip().capitalize()
secondo_nome    = st.text_input("Secondo Nome").strip().capitalize()
cf              = st.text_input("Codice Fiscale", "").strip()
telefono        = st.text_input("Mobile", "").replace(" ", "")
description     = st.text_input("PC", description_default).strip()
exp_date        = st.text_input("Data di Fine (gg-mm-aaaa)", defaults.get("expire_default", "30-06-2025")).strip()

email_flag      = st.radio("Email Consip necessaria?", ["Sì", "No"]) == "Sì"
custom_email    = ""
if not email_flag:
    custom_email = st.text_input("Email Personalizzata", "").strip()

profilazione_flag = False
sm_lines          = []
if email_flag:
    profilazione_flag = st.checkbox("Profilazione SM?")
    if profilazione_flag:
        sm_lines = st.text_area("SM su quali va profilato", "").splitlines()

# ------------------------------------------------------------
# Anteprima Messaggio (invariata)
# ------------------------------------------------------------
if email_flag and st.button("Template per Posta Elettronica"):
    sAM     = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn      = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt = formatta_data(exp_date)
    upn     = f"{sAM}@consip.it"
    mail    = upn

    st.markdown("Ciao.  \nRichiedo cortesemente la definizione di una casella di posta come sottoindicato.")
    st.markdown(f"""
| Campo             | Valore                                     |
|-------------------|--------------------------------------------|
| Tipo Utenza       | Remota                                     |
| Utenza            | {sAM}                                      |
| Alias             | {sAM}                                      |
| Display name      | {cn}                                       |
| Common name       | {cn}                                       |
| e-mail            | {mail}                                     |
| e-mail secondaria | {sAM}@consipspa.mail.onmicrosoft.com      |
"""
    )
    st.markdown("Inviare batch di notifica migrazione mail a: imac@consip.it")
    st.markdown("Aggiungere utenza di dominio ai gruppi:")
    st.markdown(f"- {o365_std}")
    st.markdown(f"- {o365_team}")
    st.markdown(f"- {o365_cop}")
    if profilazione_flag and sm_lines:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            if sm.strip(): st.markdown(f"- {sm}")
    st.markdown("Grazie  \nSaluti")

# ------------------------------------------------------------
# Unica Generazione CSV Consulente
# ------------------------------------------------------------
if st.button("Genera CSV Consulente"):
    sAM      = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn       = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt  = formatta_data(exp_date)
    upn      = f"{sAM}@consip.it"
    mail     = upn if email_flag else (custom_email or upn)
    mobile   = f"+39 {telefono}" if telefono else ""
    given    = f"{nome} {secondo_nome}".strip()
    surn     = f"{cognome} {secondo_cognome}".strip()
    inser_grp= inserimento_base if email_flag else inserimento_noemail

    # Basename normalizzato
    nc = normalize_name(cognome)
    ns = normalize_name(secondo_cognome) if secondo_cognome else ""
    basename = "_".join([nc] + ([ns] if ns else []) + [nome[:1].lower()])

    # Costruzione riga CSV
    row = [
        sAM, "SI", ou_value, cn, cn, cn, given, surn,
        cf, "", department_default, description, "No", exp_fmt,
        upn, mail, mobile, "", inser_grp, "", "", "", company
    ]
    for i in (2,3,4,5): row[i] = f"\"{row[i]}\""
    if secondo_nome:    row[6]  = f"\"{row[6]}\""
    if secondo_cognome: row[7]  = f"\"{row[7]}\""
    row[13] = f"\"{row[13]}\""
    row[16] = f"\"{row[16]}\""

    # Anteprima
    st.subheader("Anteprima CSV Consulente")
    st.dataframe(pd.DataFrame([row], columns=HEADER))

    # Download
    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
    writer.writerow(HEADER)
    writer.writerow(row)
    buf.seek(0)

    st.download_button(
        "📥 Scarica CSV Consulente",
        data=buf.getvalue(), file_name=f"{basename}_consulente.csv", mime="text/csv"
    )
    st.success(f"✅ CSV generato per '{sAM}'")
