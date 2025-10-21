import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import tempfile
from fpdf import FPDF
from docx import Document

st.title("Dashboard Marketplace & Incubateur")

# --- Fonction utilitaire ultra-robuste ---
def read_file_safe(uploaded_file, expected_columns=None):
    if uploaded_file is None or uploaded_file.size == 0:
        st.error(f"Le fichier {uploaded_file.name if uploaded_file else 'inconnu'} est vide !")
        return pd.DataFrame()

    content = uploaded_file.read()
    buffer = io.BytesIO(content)
    df = None

    if uploaded_file.name.endswith(".csv"):
        encodings = ["utf-8", "utf-8-sig", "ISO-8859-1"]
        for enc in encodings:
            try:
                buffer.seek(0)
                # Lecture CSV en devinant le séparateur
                df = pd.read_csv(buffer, encoding=enc, engine="python", sep=None, skip_blank_lines=True)
                break
            except Exception:
                df = None
        if df is None:
            st.error(f"Impossible de lire le fichier {uploaded_file.name} avec tous les encodages")
            return pd.DataFrame()
    elif uploaded_file.name.endswith(".xlsx"):
        buffer.seek(0)
        try:
            df = pd.read_excel(buffer)
        except Exception as e:
            st.error(f"Impossible de lire le fichier {uploaded_file.name} : {e}")
            return pd.DataFrame()
    else:
        st.error(f"Format de fichier non supporté : {uploaded_file.name}")
        return pd.DataFrame()

    df.columns = df.columns.str.strip()
    if expected_columns:
        missing_cols = [c for c in expected_columns if c not in df.columns]
        if missing_cols:
            st.warning(f"Colonnes manquantes dans {uploaded_file.name} : {missing_cols}")
    
    return df

# --- Upload des fichiers ---
st.sidebar.header("Uploader les fichiers")
file_users = st.sidebar.file_uploader("Noms persos", type=["csv","xlsx"])
file_entreprises = st.sidebar.file_uploader("Entreprises données", type=["csv","xlsx"])
file_mises_relation = st.sidebar.file_uploader("Historique des mises en relation", type=["csv","xlsx"])
file_base_globale = st.sidebar.file_uploader("Base globale projet", type=["csv","xlsx"])

# --- Colonnes attendues ---
cols_users = ["#Id", "Prénom", "Nom", "Inscrit depuis le", "Statut", "ID Unique", "Date de dernière connexion"]
cols_entreprises = ["Id", "Nom", "Date de création", "Date d'ouverture", "Incubateurs", "À propos", "Missions",
                    "Adresse", "Ville", "Code postal", "Téléphone", "Email", "Effectifs", "Linkedin", "Site web",
                    "Équipe", "Statut"]
cols_mises = ["Utilisateur","goBetween","Statut des mises en relation à date","Dates simples",
              "Demande de mise en relation","RDV réalisés","Taux de conversion goBetween",
              "Taux de conversion RDV réalisé","Go between validé","Go between refusé","Rdv non réalisé"]
cols_globale = ["Name","Nom","Projet","CAR/SUM (territorial)","Incubateur territorial","Statut d'incubation",
                "Poste et/ou fonction","Profil personnel Le Club","Profil sociétés Le Club",
                "Partenaires Marketplace","Date dernière connexion Le Club"]

# --- Lecture sécurisée ---
df_users = read_file_safe(file_users, expected_columns=cols_users)
df_entreprises = read_file_safe(file_entreprises, expected_columns=cols_entreprises)
df_mises = read_file_safe(file_mises_relation, expected_columns=cols_mises)
df_globale = read_file_safe(file_base_globale, expected_columns=cols_globale)

# --- Vérification que tous les fichiers sont valides ---
if not df_users.empty and not df_entreprises.empty and not df_mises.empty and not df_globale.empty:

    # --- Datas globales ---
    st.header("Datas globales")
    demandes_total = len(df_mises)
    profils_total = len(df_users)
    
    today = datetime.today()
    month_ago = today - timedelta(days=30)
    df_users["Date de dernière connexion"] = pd.to_datetime(df_users.get("Date de dernière connexion", pd.Series()), dayfirst=True, errors="coerce")
    profils_connectes = df_users[df_users["Date de dernière connexion"] >= month_ago].shape[0]
    
    st.metric("Demandes de mise en relation", demandes_total)
    st.metric("Profils créés", profils_total)
    st.metric("Profils connectés sur le mois", profils_connectes)

    # --- Marketplace ---
    st.header("Marketplace")
    go_between_valides = (df_mises.get("Go between validé", pd.Series()) == "Oui").sum()
    rdv_realises = pd.to_numeric(df_mises.get("RDV réalisés", pd.Series()), errors="coerce").sum()
    rdv_non_realises = pd.to_numeric(df_mises.get("Rdv non réalisé", pd.Series()), errors="coerce").sum()
    
    st.subheader("Totaux")
    st.metric("Go Between validés", go_between_valides)
    st.metric("RDV réalisés", rdv_realises)
    st.metric("RDV non réalisés", rdv_non_realises)

    st.subheader("Statut des mises en relation")
    st.table(df_mises.get("Statut des mises en relation à date", pd.Series()).value_counts())

    taux_go_between = pd.to_numeric(df_mises.get("Taux de conversion goBetween", pd.Series()), errors="coerce").mean()
    taux_rdv = pd.to_numeric(df_mises.get("Taux de conversion RDV réalisé", pd.Series()), errors="coerce").mean()
    st.metric("Taux de conversion Go Between (%)", round(taux_go_between, 2) if pd.notna(taux_go_between) else 0)
    st.metric("Taux de conversion RDV réalisés (%)", round(taux_rdv, 2) if pd.notna(taux_rdv) else 0)

    df_mises["Dates simples"] = pd.to_datetime(df_mises.get("Dates simples", pd.Series()), format="%Y-%m", errors="coerce")
    df_mises["Trimestre"] = df_mises["Dates simples"].dt.to_period("Q")
    trimestriel = df_mises.groupby("Trimestre").size()
    st.subheader("Répartition trimestrielle des demandes")
    st.table(trimestriel)

    # --- Profils persos & Sociétés ---
    st.header("Profils persos & Sociétés")
    st.metric("Nombre total d'entrepreneurs", len(df_globale))
    st.metric("Total profils persos", len(df_users))

    st.subheader("Statut profils persos")
    st.table(df_users.get("Statut", pd.Series()).value_counts())

    st.subheader("Statut profils sociétés")
    st.table(df_entreprises.get("Statut", pd.Series()).value_counts())

    # --- Complétion des profils (Base Globale) ---
    st.header("Complétion des profils")
    st.subheader("Vue globale")
    st.table(df_globale.get("Profil personnel Le Club", pd.Series()).value_counts())
    st.table(df_globale.get("Profil sociétés Le Club", pd.Series()).value_counts())

    st.subheader("Par CAR/SUM")
    st.table(df_globale.groupby("CAR/SUM (territorial)").get("Profil sociétés Le Club", pd.Series()).value_counts())

    st.subheader("Par Incubateur territorial")
    st.table(df_globale.groupby("Incubateur territorial").get("Profil sociétés Le Club", pd.Series()).value_counts())

    st.subheader("% de complétion sur les profils incubation individuelle")
    incubation_indiv = df_globale[df_globale.get("Statut d'incubation", pd.Series()) == "Incubation individuelle"]
    st.table(
        incubation_indiv.get("Profil sociétés Le Club", pd.Series())
        .value_counts(normalize=True)
        .mul(100)
        .round(2)
    )

    # --- Génération PDF/DOCX ---
    def generate_pdf(df_users, df_entreprises, df_mises, df_globale):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "Dashboard Marketplace & Incubateur", ln=True, align='C')
        pdf.set_font("Arial", '', 12)
        pdf.ln(5)
        pdf.cell(0, 8, f"Demandes de mise en relation: {len(df_mises)}", ln=True)
        pdf.cell(0, 8, f"Profils créés: {len(df_users)}", ln=True)
        pdf.ln(5)
        pdf.cell(0, 8, "Statut des mises en relation:", ln=True)
        for k,v in df_mises.get("Statut des mises en relation à date", pd.Series()).value_counts().items():
            pdf.cell(0, 8, f"{k}: {v}", ln=True)
        pdf.ln(5)
        pdf.cell(0, 8, "Statut profils persos:", ln=True)
        for k,v in df_users.get("Statut", pd.Series()).value_counts().items():
            pdf.cell(0, 8, f"{k}: {v}", ln=True)
        return pdf

    def generate_docx(df_users, df_entreprises, df_mises, df_globale):
        doc = Document()
        doc.add_heading("Dashboard Marketplace & Incubateur", 0)
        doc.add_heading("Datas globales", level=1)
        doc.add_paragraph(f"Demandes de mise en relation: {len(df_mises)}")
        doc.add_paragraph(f"Profils créés: {len(df_users)}")
        doc.add_heading("Statut des mises en relation", level=2)
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Statut'
        hdr_cells[1].text = 'Nombre'
        for k,v in df_mises.get("Statut des mises en relation à date", pd.Series()).value_counts().items():
            row_cells = table.add_row().cells
            row_cells[0].text = str(k)
            row_cells[1].text = str(v)
        doc.add_heading("Statut profils persos", level=2)
        table2 = doc.add_table(rows=1, cols=2)
        hdr_cells2 = table2.rows[0].cells
        hdr_cells2[0].text = 'Statut'
        hdr_cells2[1].text = 'Nombre'
        for k,v in df_users.get("Statut", pd.Series()).value_counts().items():
            row_cells = table2.add_row().cells
            row_cells[0].text = str(k)
            row_cells[1].text = str(v)
        return doc

    # --- Boutons téléchargement ---
    pdf = generate_pdf(df_users, df_entreprises, df_mises, df_globale)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        pdf.output(tmp_pdf.name)
        st.download_button("Télécharger le rapport PDF", tmp_pdf.name,
                           file_name="Dashboard_KPIs.pdf", mime="application/pdf")

    doc = generate_docx(df_users, df_entreprises, df_mises, df_globale)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_doc:
        doc.save(tmp_doc.name)
        st.download_button("Télécharger le rapport Word", tmp_doc.name,
                           file_name="Dashboard_KPIs.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

else:
    st.info("Veuillez uploader tous les fichiers correctement pour générer les KPIs.")
