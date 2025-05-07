import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------

def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name=None)
    ou_df = cfg.get("OU", pd.DataFrame(columns=["key", "label"]))
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))
    grp_df = cfg.get("InserimentoGruppi", pd.DataFrame(columns=["app", "gruppi"]))
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))
    def_df = cfg.get("Defaults", pd.DataFrame(columns=["key", "value"]))
    defaults = dict(zip(def_df["key"], def_df["value"]))
    return ou_options, gruppi, defaults

# File uploader per configurazione
st.title("Configurazione Applicazioni Streamlit")
config_file = st.file_uploader(
    "Carica config.xlsx",
    type=["xlsx"], help="File con fogli OU, InserimentoGruppi e Defaults"
)
if not config_file:
    st.warning("Carica il file di configurazione per procedere.")
    st.stop()
ou_options, gruppi, defaults = load_config_from_bytes(config_file.read())

# Utility comuni ------------------------------------------------
def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

def genera_samaccountname(nome: str, cognome: str, secondo_nome: str = "", secondo_cognome: str = "", esterno: bool = False) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20
    cand = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    cand = f"{(n[:1])}{(sn[:1])}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}")[:limit] + suffix

def build_full_name(cognome: str, secondo_cognome: str, nome: str, secondo_nome: str, esterno: bool = False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

# ------------------------------------------------------------
# App 1.3: Risorsa Esterna: Consulente
# ------------------------------------------------------------
st.header("1.3 Risorsa Esterna: Consulente")
nome_cons            = st.text_input("Nome Consulente").strip().capitalize()
secondo_nome_cons    = st.text_input("Secondo Nome").strip().capitalize()
cognome_cons         = st.text_input("Cognome").strip().capitalize()
secondo_cognome_cons = st.text_input("Secondo Cognome").strip().capitalize()
num_tel_cons         = st.text_input("Numero di Telefono", "").replace(" ", "")
desc_cons            = st.text_input("Description (lascia vuoto per <PC>)", "<PC>").strip()
cf_cons              = st.text_input("Codice Fiscale", "").strip()
exp_cons             = st.text_input("Data di Fine (gg-mm-aaaa)", defaults.get("expire_default", "30-06-2025")).strip()

ou_value_cons        = ou_options.get("esterna_consulente", "Utenti esterni - Consulenti")
employee_id_cons     = st.text_input("Employee ID", defaults.get("employee_id_default", "")).strip()
department_cons      = defaults.get("department_consulente", "Utente esterno")
inserimento_cons     = gruppi.get("esterna_consulente", "")
company_cons         = defaults.get("company_default", "")

if st.button("Genera CSV Consulente"):
    sAM = genera_samaccountname(nome_cons, cognome_cons, secondo_nome_cons, secondo_cognome_cons, True)
    cn  = build_full_name(cognome_cons, secondo_cognome_cons, nome_cons, secondo_nome_cons, True)
    exp = formatta_data(exp_cons)
    mail = f"{sAM}@consip.it"
    row = [
        sAM, "SI", ou_value_cons, cn.replace(" (esterno)", ""), cn, cn,
        " ".join([nome_cons, secondo_nome_cons]).strip(),
        " ".join([cognome_cons, secondo_cognome_cons]).strip(),
        cf_cons, employee_id_cons, department_cons,
        desc_cons or "<PC>", "No", exp,
        f"{sAM}@consip.it", mail,
        f"+39 {num_tel_cons}" if num_tel_cons else "",
        "", inserimento_cons, "", "",
        num_tel_cons, company_cons
    ]
    buf = io.StringIO(); csv.writer(buf).writerow(HEADER); csv.writer(buf).writerow(row); buf.seek(0)
    st.dataframe(pd.DataFrame([row], columns=HEADER))
    st.download_button("📥 Scarica CSV Consulente", buf.getvalue(), file_name=f"{cognome_cons}_{nome_cons[:1]}_consulente.csv", mime="text/csv")
    st.success(f"✅ Generato {sAM}")
