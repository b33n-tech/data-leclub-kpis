import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.title("Dashboard Marketplace & Incubateur")

# --- Fonction utilitaire pour lire CSV/XLSX robustement ---
def read_file_safe(uploaded_file, expected_columns=None):
    """
    Lit un fichier CSV ou XLSX uploadé, gère encodages, fichiers vides et colonnes manquantes.
    """
    if uploaded_file is None or uploaded_file.size == 0:
        st.error(f"Le fichier {uploaded_file.name if uploaded_file else 'inconnu'} est vide !")
        return pd.DataFrame()

    content = uploaded_file.read()
    buffer = io.BytesIO(content)

    try:
        if uploaded_file.name.endswith(".csv"):
            try:
                df = pd.read_csv(buffer, encoding="utf-8")
            except UnicodeDecodeError:
                try:
                    buffer.seek(0)
                    df = pd.read_csv(buffer, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    buffer.seek(0)
                    df = pd.read_csv(buffer, encoding="ISO-8859-1", engine="python", errors="replace")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(buffer)
        else:
            st.error(f"Format de fichier non supporté : {uploaded_file.name}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Impossible de lire le fichier {uploaded_file.name} : {e}")
        return pd.DataFrame()

    # Nettoyage colonnes
    df.columns = df.columns.str.strip()

    # Vérification colonnes attendues
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

else:
    st.info("Veuillez uploader tous les fichiers correctement pour générer les KPIs.")
