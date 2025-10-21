import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import plotly.express as px

st.title("Dashboard Marketplace & Incubateur (V2 Interactive & Robuste)")

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

    # Nettoyage colonnes
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

# --- Vérification ---
if not df_users.empty and not df_entreprises.empty and not df_mises.empty and not df_globale.empty:

    # --- KPIs globaux ---
    st.header("KPIs Globaux")
    demandes_total = len(df_mises)
    profils_total = len(df_users)
    
    today = datetime.today()
    month_ago = today - timedelta(days=30)
    df_users["Date de dernière connexion"] = pd.to_datetime(df_users["Date de dernière connexion"], dayfirst=True, errors="coerce")
    profils_connectes = df_users[df_users["Date de dernière connexion"] >= month_ago].shape[0]

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Demandes de mise en relation", demandes_total)
    kpi2.metric("Profils créés", profils_total)
    kpi3.metric("Profils connectés sur le mois", profils_connectes)

    # --- Marketplace ---
    st.header("Marketplace")
    
    # Dates & Trimestre
    if "Dates simples" in df_mises.columns:
        df_mises["Dates simples"] = pd.to_datetime(df_mises["Dates simples"], format="%Y-%m", errors="coerce")
        df_mises["Trimestre"] = df_mises["Dates simples"].dt.to_period("Q").astype(str)
    else:
        df_mises["Trimestre"] = pd.Series(dtype=str)

    # Statut mises en relation
    if "Statut des mises en relation à date" in df_mises.columns:
        fig_statut = px.histogram(
            df_mises,
            x="Statut des mises en relation à date",
            title="Répartition des statuts des mises en relation",
            labels={"count": "Nombre"},
            text_auto=True
        )
        st.plotly_chart(fig_statut, use_container_width=True)

    # Répartition trimestrielle
    if not df_mises["Trimestre"].empty:
        trimestriel = df_mises.groupby("Trimestre").size().reset_index(name="Nombre de demandes")
        fig_trimestriel = px.bar(
            trimestriel,
            x="Trimestre",
            y="Nombre de demandes",
            title="Répartition trimestrielle des demandes",
            text="Nombre de demandes"
        )
        st.plotly_chart(fig_trimestriel, use_container_width=True)

    # Taux de conversion
    taux_df = pd.DataFrame({
        "Indicateur": ["Go Between", "RDV réalisés"],
        "Taux (%)": [
            pd.to_numeric(df_mises.get("Taux de conversion goBetween", pd.Series([0])), errors="coerce").mean(),
            pd.to_numeric(df_mises.get("Taux de conversion RDV réalisé", pd.Series([0])), errors="coerce").mean()
        ]
    })
    fig_taux = px.bar(
        taux_df,
        x="Indicateur",
        y="Taux (%)",
        title="Taux de conversion moyen",
        text="Taux (%)"
    )
    st.plotly_chart(fig_taux, use_container_width=True)

    # --- Profils persos & Sociétés ---
    st.header("Profils persos & Sociétés")
    
    if "Statut" in df_users.columns:
        fig_users_statut = px.pie(
            df_users,
            names="Statut",
            title="Répartition des statuts profils persos"
        )
        st.plotly_chart(fig_users_statut, use_container_width=True)

    if "Statut" in df_entreprises.columns:
        fig_entreprises_statut = px.pie(
            df_entreprises,
            names="Statut",
            title="Répartition des statuts profils sociétés"
        )
        st.plotly_chart(fig_entreprises_statut, use_container_width=True)

    # --- Complétion des profils ---
    st.header("Complétion des profils (Base Globale)")

    # Profils personnels
    if "Profil personnel Le Club" in df_globale.columns:
        df_globale["Profil personnel Le Club"] = df_globale["Profil personnel Le Club"].astype(str).str.strip()
        persos_count = df_globale["Profil personnel Le Club"].value_counts().reset_index()
        persos_count.columns = ["Statut", "Nombre"]
        if not persos_count.empty:
            fig_complet_persos = px.bar(
                persos_count,
                x="Statut",
                y="Nombre",
                title="Complétion profils personnels",
                text="Nombre"
            )
            st.plotly_chart(fig_complet_persos, use_container_width=True)
        else:
            st.info("Aucun profil personnel à afficher.")

    # Profils sociétés
    if "Profil sociétés Le Club" in df_globale.columns:
        df_globale["Profil sociétés Le Club"] = df_globale["Profil sociétés Le Club"].astype(str).str.strip()
        societes_count = df_globale["Profil sociétés Le Club"].value_counts().reset_index()
        societes_count.columns = ["Statut", "Nombre"]
        if not societes_count.empty:
            fig_complet_societes = px.bar(
                societes_count,
                x="Statut",
                y="Nombre",
                title="Complétion profils sociétés",
                text="Nombre"
            )
            st.plotly_chart(fig_complet_societes, use_container_width=True)
        else:
            st.info("Aucun profil société à afficher.")

else:
    st.info("Veuillez uploader tous les fichiers correctement pour générer les KPIs.")
