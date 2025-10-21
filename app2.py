import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
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
                df = pd.read_csv(buffer, encoding=enc, engine="python")
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
            st.error(f"Colonnes manquantes dans {uploaded_file.name} : {missing_cols}")

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
    df_users["Date de dernière connexion"] = pd.to_datetime(df_users["Date de dernière connexion"], dayfirst=True, errors="coerce")
    profils_connectes = df_users[df_users["Date de dernière connexion"] >= month_ago].shape[0]
    
    st.metric("Demandes de mise en relation", demandes_total)
    st.metric("Profils créés", profils_total)
    st.metric("Profils connectés sur le mois", profils_connectes)

    # --- Marketplace ---
    st.header("Marketplace")
    go_between_valides = (df_mises["Go between validé"] == "Oui").sum()
    rdv_realises = df_mises["RDV réalisés"].sum()
    rdv_non_realises = df_mises["Rdv non réalisé"].sum()
    
    st.subheader("Totaux")
    st.metric("Go Between validés", go_between_valides)
    st.metric("RDV réalisés", rdv_realises)
    st.metric("RDV non réalisés", rdv_non_realises)

    st.subheader("Statut des mises en relation")
    st.table(df_mises["Statut des mises en relation à date"].value_counts())

    taux_go_between = pd.to_numeric(df_mises["Taux de conversion goBetween"], errors="coerce").mean()
    taux_rdv = pd.to_numeric(df_mises["Taux de conversion RDV réalisé"], errors="coerce").mean()
    st.metric("Taux de conversion Go Between (%)", round(taux_go_between, 2))
    st.metric("Taux de conversion RDV réalisés (%)", round(taux_rdv, 2))

    df_mises["Dates simples"] = pd.to_datetime(df_mises["Dates simples"], format="%Y-%m", errors="coerce")
    df_mises["Trimestre"] = df_mises["Dates simples"].dt.to_period("Q")
    trimestriel = df_mises.groupby("Trimestre").size()
    st.subheader("Répartition trimestrielle des demandes")
    st.table(trimestriel)

    # --- Profils persos & Sociétés ---
    st.header("Profils persos & Sociétés")
    st.metric("Nombre total d'entrepreneurs", len(df_globale))
    st.metric("Total profils persos", len(df_users))

    st.subheader("Statut profils persos")
    st.table(df_users["Statut"].value_counts())

    st.subheader("Statut profils sociétés")
    st.table(df_entreprises["Statut"].value_counts())

    # --- Complétion des profils (Base Globale) ---
    st.header("Complétion des profils")
    st.subheader("Vue globale")
    st.table(df_globale["Profil personnel Le Club"].value_counts())
    st.table(df_globale["Profil sociétés Le Club"].value_counts())

    st.subheader("Par CAR/SUM")
    st.table(df_globale.groupby("CAR/SUM (territorial)")["Profil sociétés Le Club"].value_counts())

    st.subheader("Par Incubateur territorial")
    st.table(df_globale.groupby("Incubateur territorial")["Profil sociétés Le Club"].value_counts())

    st.subheader("% de complétion sur les profils incubation individuelle")
    incubation_indiv = df_globale[df_globale["Statut d'incubation"] == "Incubation individuelle"]
    st.table(
        incubation_indiv["Profil sociétés Le Club"]
        .value_counts(normalize=True)
        .mul(100)
        .round(2)
    )

    # --- Génération du DOCX ---
    def generate_docx(df_users, df_entreprises, df_mises, df_globale):
        doc = Document()
        doc.add_heading("Dashboard Marketplace & Incubateur - Extract", 0)

        # --- KPIs ---
        doc.add_heading("KPIs", level=1)
        doc.add_paragraph(f"Demandes de mise en relation : {len(df_mises)}")
        doc.add_paragraph(f"Profils créés : {len(df_users)}")
        doc.add_paragraph(f"Profils connectés sur le mois : {profils_connectes}")

        # --- Marketplace ---
        doc.add_heading("Marketplace", level=1)
        doc.add_paragraph(f"Go Between validés : {go_between_valides}")
        doc.add_paragraph(f"RDV réalisés : {rdv_realises}")
        doc.add_paragraph(f"RDV non réalisés : {rdv_non_realises}")
        doc.add_paragraph(f"Taux de conversion Go Between (%) : {round(taux_go_between,2)}")
        doc.add_paragraph(f"Taux de conversion RDV réalisés (%) : {round(taux_rdv,2)}")

        # --- Tableaux principaux ---
        def df_to_docx_table(df, title):
            doc.add_heading(title, level=2)
            table = doc.add_table(rows=1, cols=len(df.columns))
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, col in enumerate(df.columns):
                hdr_cells[i].text = str(col)
            for _, row in df.iterrows():
                row_cells = table.add_row().cells
                for i, val in enumerate(row):
                    row_cells[i].text = str(val)
        
        df_to_docx_table(df_users.head(10), "Exemple Profils Persos (10 lignes)")
        df_to_docx_table(df_entreprises.head(10), "Exemple Profils Sociétés (10 lignes)")
        df_to_docx_table(df_mises.head(10), "Exemple Mises en Relation (10 lignes)")
        df_to_docx_table(df_globale.head(10), "Exemple Base Globale (10 lignes)")

        doc_stream = io.BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        return doc_stream

    docx_data = generate_docx(df_users, df_entreprises, df_mises, df_globale)
    st.download_button(
        label="Télécharger l'extract en DOCX",
        data=docx_data,
        file_name=f"dashboard_extract_{datetime.today().strftime('%Y%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

else:
    st.info("Veuillez uploader tous les fichiers correctement pour générer les KPIs.")
