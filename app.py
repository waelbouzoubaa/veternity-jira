import os
import pandas as pd
import streamlit as st

# "google_sheets" (prod / test VPS, portable) ou "excel" (test local rapide, voir generate_sample.py)
DATA_SOURCE = os.getenv("DATA_SOURCE", "excel")
EXCEL_PATH = os.getenv("EXCEL_OUTPUT_PATH", "tickets_enrichis.xlsx")
SHEET_NAME = "Tickets enrichis"

GOOGLE_SHEETS_CREDENTIALS_PATH = os.getenv(
    "GOOGLE_SHEETS_CREDENTIALS_PATH", "/app/credentials/ai-portal-bq-analytics-key.json"
)
GOOGLE_SHEETS_SPREADSHEET_ID = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "")
GOOGLE_SHEETS_TAB_NAME = os.getenv("GOOGLE_SHEETS_TAB_NAME", "")

st.set_page_config(page_title="Tickets Jira enrichis", layout="wide")
st.title("🎫 Tickets Jira enrichis par le RAG Confluence")


def parse_score(value) -> float:
    """Le Google Sheet peut stocker le score en pourcentage affiché ("72%")
    plutôt qu'en nombre brut (0.72) selon le format de cellule appliqué."""
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", ".")
    if s.endswith("%"):
        s = s[:-1]
        try:
            return float(s) / 100
        except ValueError:
            return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


@st.cache_resource
def _get_worksheet():
    import gspread
    from google.oauth2.service_account import Credentials as GoogleCredentials

    creds = GoogleCredentials.from_service_account_file(
        GOOGLE_SHEETS_CREDENTIALS_PATH,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    client = gspread.authorize(creds)
    sh = client.open_by_key(GOOGLE_SHEETS_SPREADSHEET_ID)
    return sh.worksheet(GOOGLE_SHEETS_TAB_NAME) if GOOGLE_SHEETS_TAB_NAME else sh.sheet1


@st.cache_data(ttl=60)
def load_data() -> pd.DataFrame:
    if DATA_SOURCE == "excel":
        if not os.path.exists(EXCEL_PATH):
            return pd.DataFrame()
        return pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    ws = _get_worksheet()
    records = ws.get_all_records()
    return pd.DataFrame(records)


col_refresh, col_source = st.columns([1, 5])
with col_refresh:
    if st.button("🔄 Rafraîchir"):
        load_data.clear()
        st.rerun()
with col_source:
    st.caption(f"Source : {DATA_SOURCE}")

df = load_data()

if df.empty:
    st.warning("Aucune donnée trouvée pour le moment.")
    st.stop()

df["Score de pertinence"] = df["Score de pertinence"].apply(parse_score)

# --- Filtres ---
c1, c2, c3 = st.columns(3)
with c1:
    statuts = st.multiselect("Statut", sorted(df["Statut"].dropna().unique()))
with c2:
    score_min = st.slider("Score minimum", 0.0, 1.0, 0.0, 0.05)
with c3:
    search = st.text_input("Recherche (ticket, titre)")

filtered = df.copy()
if statuts:
    filtered = filtered[filtered["Statut"].isin(statuts)]
filtered = filtered[filtered["Score de pertinence"] >= score_min]
if search:
    mask = (
        filtered["Ticket"].astype(str).str.contains(search, case=False, na=False)
        | filtered["Titre"].astype(str).str.contains(search, case=False, na=False)
    )
    filtered = filtered[mask]

filtered = filtered.sort_values("Date de création", ascending=False)

PAGE_SIZE = 30
total = len(filtered)
total_pages = max(1, (total - 1) // PAGE_SIZE + 1)
page = st.number_input("Page", min_value=1, max_value=total_pages, value=1, step=1)
start = (page - 1) * PAGE_SIZE
end = start + PAGE_SIZE
page_df = filtered.iloc[start:end]

st.caption(f"{total} ticket(s) au total sur {len(df)} — page {page}/{total_pages} (affichage {start + 1}-{min(end, total)})")

# --- Liste des tickets (page courante uniquement) ---
for _, row in page_df.iterrows():
    score = row.get("Score de pertinence") or 0.0
    badge = "🟢" if score >= 0.7 else ("🟡" if score >= 0.5 else "🔴")
    ticket_key = str(row["Ticket"])

    with st.container(border=True):
        left, right = st.columns([4, 1])
        with left:
            st.markdown(f"**{ticket_key}** — {row['Titre']}")
            st.caption(f"Statut : {row['Statut']} · Créé le {row['Date de création']}")
            if row.get("Lien Jira"):
                st.link_button("🔗 Ouvrir le ticket Jira", row["Lien Jira"])
        with right:
            st.markdown(f"### {badge} {score:.2f}")

        with st.expander("Résumé RAG"):
            st.write(row.get("Résumé RAG") or "—")

        ressources = str(row.get("Ressources trouvées") or "")
        if ressources and ressources != "Aucune ressource pertinente trouvée":
            st.markdown("**Sources Confluence consultées :**")
            lines = [l for l in ressources.split("\n") if l.strip()]
            btn_cols = st.columns(min(len(lines), 3))
            for i, line in enumerate(lines):
                if " — " in line:
                    title, url = line.rsplit(" — ", 1)
                else:
                    title, url = line, None
                with btn_cols[i % len(btn_cols)]:
                    if url:
                        st.link_button(f"📄 {title[:45]}", url)
                    else:
                        st.caption(title)
        else:
            st.caption("Aucune ressource pertinente trouvée")
