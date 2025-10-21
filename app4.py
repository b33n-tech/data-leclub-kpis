import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

st.title("Dashboard Marketplace & Incubateur (V8 Interactive)")

# --- Fonction utilitaire ---
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

# --- Upload fichiers ---
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

# --- Vérification fichiers ---
if not df_users.empty and not df_entreprises.empty and not df_mises.empty and not df_globale.empty:

    # --- Nettoyage ---
    df_users["Date de dernière connexion"] = pd.to_datetime(df_users["Date de dernière connexion"], dayfirst=True, errors="coerce")
    df_mises["Dates simples"] = pd.to_datetime(df_mises.get("Dates simples", pd.Series([])), format="%Y-%m", errors="coerce")
    df_mises["Trimestre"] = df_mises["Dates simples"].dt.to_period("Q").astype(str)
    df_globale["Profil personnel Le Club"] = df_globale.get("Profil personnel Le Club", pd.Series([])).astype(str).str.strip()
    df_globale["Profil sociétés Le Club"] = df_globale.get("Profil sociétés Le Club", pd.Series([])).astype(str).str.strip()

    # --- Tableau interactif avec st_aggrid ---
    st.subheader("Tableau interactif (filtrez les colonnes pour recalculer KPIs)")
    gb = GridOptionsBuilder.from_dataframe(df_globale)
    gb.configure_default_column(filterable=True, sortable=True)
    gb.configure_selection(selection_mode="multiple", use_checkbox=True)
    gridOptions = gb.build()

    grid_response = AgGrid(
        df_globale,
        gridOptions=gridOptions,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        fit_columns_on_grid_load=True,
        height=400,
        reload_data=True
    )

    # --- Récupérer les lignes visibles après filtrage ---
    df_filtered = pd.DataFrame(grid_response['data'])

    # --- KPIs recalculés dynamiques ---
    st.subheader("KPIs dynamiques selon filtre")
    demandes_total = df_mises[df_mises["Utilisateur"].isin(df_filtered["Name"])].shape[0]
    profils_total = len(df_filtered)
    today = datetime.today()
    month_ago = today - timedelta(days=30)
    profils_connectes = df_users[df_users["Date de dernière connexion"] >= month_ago].shape[0]

    col1, col2, col3 = st.columns(3)
    col1.metric("Profils filtrés", profils_total)
    col2.metric("Demandes Marketplace filtrées", demandes_total)
    col3.metric("Profils connectés ce mois", profils_connectes)

    # --- Graphiques dynamiques ---
    st.subheader("Statut profils personnels")
    persos_count = df_filtered["Profil personnel Le Club"].value_counts()
    if not persos_count.empty:
        fig_persos = px.bar(persos_count.reset_index(), x="index", y="Profil personnel Le Club", text="Profil personnel Le Club")
        fig_persos.update_layout(xaxis_title="Profil personnel", yaxis_title="Nombre")
        st.plotly_chart(fig_persos, use_container_width=True)

    st.subheader("Statut profils sociétés")
    societes_count = df_filtered["Profil sociétés Le Club"].value_counts()
    if not societes_count.empty:
        fig_soc = px.bar(societes_count.reset_index(), x="index", y="Profil sociétés Le Club", text="Profil sociétés Le Club")
        fig_soc.update_layout(xaxis_title="Profil sociétés", yaxis_title="Nombre")
        st.plotly_chart(fig_soc, use_container_width=True)

    st.subheader("Marketplace (Demandes par statut)")
    df_mises_filtered = df_mises[df_mises["Utilisateur"].isin(df_filtered["Name"])]
    if not df_mises_filtered.empty:
        status_counts = df_mises_filtered["Statut des mises en relation à date"].value_counts()
        fig_market = px.bar(status_counts.reset_index(), x="index", y="Statut des mises en relation à date", text="Statut des mises en relation à date")
        fig_market.update_layout(xaxis_title="Statut", yaxis_title="Nombre de demandes")
        st.plotly_chart(fig_market, use_container_width=True)

else:
    st.info("Veuillez uploader tous les fichiers correctement pour générer les KPIs.")
