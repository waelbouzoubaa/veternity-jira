import html
import os
import re
import unicodedata

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# "google_sheets" (prod / test VPS, portable) ou "excel" (test local rapide, voir generate_sample.py)
DATA_SOURCE = os.getenv("DATA_SOURCE", "excel")
EXCEL_PATH = os.getenv("EXCEL_OUTPUT_PATH", "tickets_enrichis.xlsx")
SHEET_NAME = "Tickets enrichis"

GOOGLE_SHEETS_CREDENTIALS_PATH = os.getenv(
    "GOOGLE_SHEETS_CREDENTIALS_PATH", "/app/credentials/ai-portal-bq-analytics-key.json"
)
GOOGLE_SHEETS_SPREADSHEET_ID = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "")
GOOGLE_SHEETS_TAB_NAME = os.getenv("GOOGLE_SHEETS_TAB_NAME", "")

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.png")

st.set_page_config(page_title="Tickets Jira enrichis", page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "🎫", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1100px; }
      p, .stMarkdown { line-height: 1.55; }
    </style>
    """,
    unsafe_allow_html=True,
)

header_logo, header_title = st.columns([1, 5])
with header_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=140)
with header_title:
    st.markdown(
        "<h2 style='margin-top:1.2rem;'>Tickets Jira enrichis par le RAG Confluence</h2>",
        unsafe_allow_html=True,
    )
st.divider()


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


def strip_formatting(text: str) -> str:
    """Le résumé RAG contient des balises <br> et du **gras** markdown (format imposé
    par le prompt du RAG Confluence) — on les retire pour un aperçu en texte propre,
    tronquable sans risque de couper une balise en plein milieu."""
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"\s+", " ", text).strip()
    return html.escape(text)


def _normalize_col(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def find_column(df: pd.DataFrame, target: str):
    """Retrouve la colonne correspondant à `target` même si le Sheet a une
    orthographe légèrement différente (accents, espaces, majuscules)."""
    target_norm = _normalize_col(target)
    for col in df.columns:
        if _normalize_col(col) == target_norm:
            return col
    return None


def score_badge(score: float) -> str:
    if score >= 0.7:
        bg = "#2e7d32"
    elif score >= 0.5:
        bg = "#b8860b"
    else:
        bg = "#c62828"
    return (
        f"<span style='background:{bg};color:#ffffff;padding:4px 12px;"
        f"border-radius:14px;font-weight:600;font-size:0.9rem;white-space:nowrap;'>"
        f"{score:.2f}</span>"
    )


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

with st.expander("🔧 Diagnostic colonnes (colonnes détectées dans la source)"):
    st.write(list(df.columns))

_comment_col_found = find_column(df, "Commentaires équipe support")
if _comment_col_found is None:
    COMMENT_COL = "Commentaires équipe support"
    df[COMMENT_COL] = ""
    st.warning(
        "⚠️ Colonne 'Commentaires équipe support' introuvable dans la source — "
        "vérifie l'intitulé exact dans le diagnostic ci-dessus."
    )
else:
    COMMENT_COL = _comment_col_found

# --- Filtres ---
c1, c2, c3, c4 = st.columns([2, 2, 3, 2])
with c1:
    statuts = st.multiselect("Statut", sorted(df["Statut"].dropna().unique()))
with c2:
    score_min = st.slider("Score minimum", 0.0, 1.0, 0.0, 0.05)
with c3:
    search = st.text_input("Recherche (ticket, titre)")
with c4:
    only_commented = st.checkbox("💬 Avec commentaire support")

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
if only_commented:
    filtered = filtered[filtered[COMMENT_COL].apply(lambda v: not pd.isna(v) and str(v).strip() != "")]

filtered = filtered.sort_values("Date de création", ascending=False)

PAGE_SIZE = 10
total = len(filtered)
total_pages = max(1, (total - 1) // PAGE_SIZE + 1)

if "page" not in st.session_state:
    st.session_state.page = 1
if st.session_state.page > total_pages:
    st.session_state.page = 1
page = st.session_state.page


def _pages_to_show(current: int, last: int) -> list:
    if last <= 7:
        return list(range(1, last + 1))
    keep = {1, last, current}
    for d in (-1, 1, -2, 2):
        p = current + d
        if 1 <= p <= last:
            keep.add(p)
    ordered = sorted(keep)
    result = []
    prev = None
    for p in ordered:
        if prev is not None and p - prev > 1:
            result.append("…")
        result.append(p)
        prev = p
    return result


def render_pagination(current: int, last: int, suffix: str) -> None:
    page_list = _pages_to_show(current, last)
    nav_cols = st.columns([1] * (len(page_list) + 2) + [6])

    with nav_cols[0]:
        if st.button("‹", key=f"prev_{suffix}", disabled=current <= 1, use_container_width=True):
            st.session_state.page = current - 1
            st.rerun()
    for i, p in enumerate(page_list):
        with nav_cols[i + 1]:
            if p == "…":
                st.markdown("<div style='text-align:center;padding-top:8px;'>…</div>", unsafe_allow_html=True)
            elif st.button(str(p), key=f"page_btn_{suffix}_{p}", type="primary" if p == current else "secondary", use_container_width=True):
                st.session_state.page = p
                st.rerun()
    with nav_cols[len(page_list) + 1]:
        if st.button("›", key=f"next_{suffix}", disabled=current >= last, use_container_width=True):
            st.session_state.page = current + 1
            st.rerun()


start = (page - 1) * PAGE_SIZE
end = start + PAGE_SIZE
page_df = filtered.iloc[start:end]

st.caption(f"{total} ticket(s) au total sur {len(df)} — page {page}/{total_pages} (affichage {start + 1}-{min(end, total)})")

# --- Liste des tickets (page courante uniquement) ---
PREVIEW_LEN = 200

for _, row in page_df.iterrows():
    score = row.get("Score de pertinence") or 0.0
    ticket_key = str(row["Ticket"])

    with st.container(border=True):
        head_left, head_right = st.columns([5, 1])
        with head_left:
            st.markdown(
                f"<div style='font-size:1.15rem;font-weight:700;color:#3f3f3e;'>{ticket_key}</div>"
                f"<div style='font-size:1rem;color:#3f3f3e;margin-top:2px;'>{row['Titre']}</div>",
                unsafe_allow_html=True,
            )
        with head_right:
            st.markdown(
                f"<div style='text-align:right;padding-top:6px;'>{score_badge(score)}</div>",
                unsafe_allow_html=True,
            )

        st.markdown(
            f"<div style='color:#8a8a88;font-size:0.85rem;margin:8px 0 12px 0;'>"
            f"Statut : <b>{row['Statut']}</b> &nbsp;·&nbsp; Créé le {row['Date de création']}</div>",
            unsafe_allow_html=True,
        )

        if row.get("Lien Jira"):
            st.link_button("🔗 Ouvrir le ticket Jira", row["Lien Jira"])

        comment_val = row.get(COMMENT_COL)
        comment = "" if pd.isna(comment_val) else str(comment_val).strip()
        if comment:
            st.markdown(
                f"<div style='background:#fff3cd;border-left:4px solid #b8860b;"
                f"padding:8px 12px;border-radius:4px;margin-top:10px;color:#3f3f3e;'>"
                f"💬 <b>Commentaire équipe support :</b> {html.escape(comment)}</div>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='margin-top:14px;'></div>", unsafe_allow_html=True)

        resume_raw = str(row.get("Résumé RAG") or "—")
        resume_preview = strip_formatting(resume_raw)
        st.markdown(
            f"<div style='color:#3f3f3e;'>{resume_preview[:PREVIEW_LEN].rstrip()}"
            f"{'…' if len(resume_preview) > PREVIEW_LEN else ''}</div>",
            unsafe_allow_html=True,
        )
        if len(resume_preview) > PREVIEW_LEN:
            with st.expander("Voir le résumé complet"):
                st.markdown(resume_raw, unsafe_allow_html=True)

        ressources = str(row.get("Ressources trouvées") or "")
        if ressources and ressources != "Aucune ressource pertinente trouvée":
            st.markdown(
                "<div style='margin-top:16px;font-weight:600;color:#3f3f3e;'>"
                "📚 Sources Confluence consultées</div>",
                unsafe_allow_html=True,
            )
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

    st.markdown("<div style='margin-bottom:14px;'></div>", unsafe_allow_html=True)

st.divider()
render_pagination(page, total_pages, "bottom")
