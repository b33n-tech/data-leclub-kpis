import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import plotly.graph_objects as go
import plotly.express as px

st.title("Dashboard Marketplace & Incubateur (V6 Interactive)")

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

    # -----------------------------
    # Marketplace
    st.header("Marketplace")
    st.sidebar.subheader("Filtres Marketplace")
    car_market = ["Tout"] + sorted(df_globale["CAR/SUM (territorial)"].dropna().unique())
    selected_car_market = st.sidebar.selectbox("CAR/SUM (Marketplace)", car_market, key="car_market")
    inc_market = ["Tout"] + sorted(df_globale["Incubateur territorial"].dropna().unique())
    selected_inc_market = st.sidebar.selectbox("Incubateur (Marketplace)", inc_market, key="inc_market")
    trimestre_market = ["Tout"] + sorted(df_mises["Trimestre"].dropna().unique())
    selected_trimestre_market = st.sidebar.selectbox("Trimestre (Marketplace)", trimestre_market, key="trimestre_market")

    # Filtrage Marketplace
    df_mises_market = df_mises.copy()
    df_globale_market = df_globale.copy()
    if selected_trimestre_market != "Tout":
        df_mises_market = df_mises_market[df_mises_market["Trimestre"] == selected_trimestre_market]
    if selected_car_market != "Tout":
        df_globale_market = df_globale_market[df_globale_market["CAR/SUM (territorial)"] == selected_car_market]
    if selected_inc_market != "Tout":
        df_globale_market = df_globale_market[df_globale_market["Incubateur territorial"] == selected_inc_market]

    # KPIs Marketplace
    demandes_total = len(df_mises_market)
    go_between_valides = (df_mises_market["Go between validé"] == "Oui").sum()
    rdv_realises = df_mises_market["RDV réalisés"].sum()
    rdv_non_realises = df_mises_market["Rdv non réalisé"].sum()
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Demandes", demandes_total)
    col2.metric("Go Between validés", go_between_valides)
    col3.metric("RDV réalisés", rdv_realises)
    col4.metric("RDV non réalisés", rdv_non_realises)

    # Graphique statut
    if not df_mises_market.empty:
        status_counts = df_mises_market["Statut des mises en relation à date"].value_counts()
        fig_market = go.Figure()
        for status in status_counts.index:
            fig_market.add_trace(go.Bar(x=[status], y=[status_counts[status]], name=status))
        fig_market.update_layout(
            title="Statut des mises en relation",
            updatemenus=[dict(
                buttons=[dict(label=s,
                              method="update",
                              args=[{"visible":[st==s for st in status_counts.index]}])
                         for s in status_counts.index] +
                        [dict(label="Tous", method="update", args=[{"visible":[True]*len(status_counts)}])]
            )]
        )
        st.plotly_chart(fig_market, use_container_width=True)

    # Trimestriel
    if not df_mises_market.empty:
        trimestriel = df_mises_market.groupby("Trimestre").size().reset_index(name="Nombre de demandes")
        fig_tri = px.bar(trimestriel, x="Trimestre", y="Nombre de demandes", title="Répartition trimestrielle")
        st.plotly_chart(fig_tri, use_container_width=True)

    # -----------------------------
    # Profils personnels
    st.header("Profils personnels")
    st.sidebar.subheader("Filtres Profils personnels")
    car_perso = ["Tout"] + sorted(df_globale["CAR/SUM (territorial)"].dropna().unique())
    selected_car_perso = st.sidebar.selectbox("CAR/SUM (Profils perso)", car_perso, key="car_perso")
    inc_perso = ["Tout"] + sorted(df_globale["Incubateur territorial"].dropna().unique())
    selected_inc_perso = st.sidebar.selectbox("Incubateur (Profils perso)", inc_perso, key="inc_perso")

    df_persos_filtered = df_globale.copy()
    if selected_car_perso != "Tout":
        df_persos_filtered = df_persos_filtered[df_persos_filtered["CAR/SUM (territorial)"] == selected_car_perso]
    if selected_inc_perso != "Tout":
        df_persos_filtered = df_persos_filtered[df_persos_filtered["Incubateur territorial"] == selected_inc_perso]

    if not df_persos_filtered.empty:
        persos_count = df_persos_filtered["Profil personnel Le Club"].value_counts()
        fig_persos = go.Figure()
        for st_name in persos_count.index:
            fig_persos.add_trace(go.Bar(x=[st_name], y=[persos_count[st_name]], name=st_name))
        fig_persos.update_layout(
            title="Complétion Profils personnels",
            updatemenus=[dict(
                buttons=[dict(label=n,
                              method="update",
                              args=[{"visible":[nm==n for nm in persos_count.index]}])
                         for n in persos_count.index] +
                        [dict(label="Tous", method="update", args=[{"visible":[True]*len(persos_count)}])]
            )]
        )
        st.plotly_chart(fig_persos, use_container_width=True)

    # -----------------------------
    # Profils sociétés
    st.header("Profils sociétés")
    st.sidebar.subheader("Filtres Profils sociétés")
    car_soc = ["Tout"] + sorted(df_globale["CAR/SUM (territorial)"].dropna().unique())
    selected_car_soc = st.sidebar.selectbox("CAR/SUM (Sociétés)", car_soc, key="car_soc")
    inc_soc = ["Tout"] + sorted(df_globale["Incubateur territorial"].dropna().unique())
    selected_inc_soc = st.sidebar.selectbox("Incubateur (Sociétés)", inc_soc, key="inc_soc")

    df_soc_filtered = df_globale.copy()
    if selected_car_soc != "Tout":
        df_soc_filtered = df_soc_filtered[df_soc_filtered["CAR/SUM (territorial)"] == selected_car_soc]
    if selected_inc_soc != "Tout":
        df_soc_filtered = df_soc_filtered[df_soc_filtered["Incubateur territorial"] == selected_inc_soc]

    if not df_soc_filtered.empty:
        societes_count = df_soc_filtered["Profil sociétés Le Club"].value_counts()
        fig_soc = go.Figure()
        for st_name in societes_count.index:
            fig_soc.add_trace(go.Bar(x=[st_name], y=[societes_count[st_name]], name=st_name))
        fig_soc.update_layout(
            title="Complétion Profils sociétés",
            updatemenus=[dict(
                buttons=[dict(label=n,
                              method="update",
                              args=[{"visible":[nm==n for nm in societes_count.index]}])
                         for n in societes_count.index] +
                        [dict(label="Tous", method="update", args=[{"visible":[True]*len(societes_count)}])]
            )]
        )
        st.plotly_chart(fig_soc, use_container_width=True)

else:
    st.info("Veuillez uploader tous les fichiers correctement pour générer les KPIs.")
