import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.title("Dashboard Marketplace & Incubateur")

# --- Fonction utilitaire pour lire CSV/XLSX ---
def read_file_safe(uploaded_file, expected_columns=None):
    """
    Lit un fichier CSV ou XLSX uploadé, gère les encodages, fichiers vides et vérifie les colonnes.
    """
    if uploaded_file is None:
        return pd.DataFrame()
    
    # Vérifier que le fichier n'est pas vide
    if uploaded_file.size == 0:
        st.error(f"Le fichier {uploaded_file.name} est vide !")
        return pd.DataFrame()
    
    # Lecture selon l'extension
    try:
        if uploaded_file.name.endswith(".csv"):
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8")
            except UnicodeDecodeError:
                try:
                    df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    df = pd.read_csv(uploaded_file, encoding="ISO-8859-1", errors="replace")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file)
        else:
            st.error(f"Format de fichier non supporté : {uploaded_file.name}")
            return pd.DataFrame()
    except pd.errors.EmptyDataError:
        st.error(f"Le fichier {uploaded_file.name} est vide ou illisible !")
        return pd.DataFrame()
    
    # Nettoyer les colonnes
    df.columns = df.columns.str.strip()
    
    # Vérifier colonnes attendues
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

# Colonnes attendues (exemple)
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

# --- Lecture sécurisée des fichiers ---
df_users = read_file_safe(file_users, expected_columns=cols_users)
df_entreprises = read_file_safe(file_entreprises, expected_columns=cols_entreprises)
df_mises = read_file_safe(file_mises_relation, expected_columns=cols_mises)
df_globale = read_file_safe(file_base_globale, expected_columns=cols_globale)

# --- Vérification que tous les fichiers sont valides avant calcul KPIs ---
if not df_users.empty and not df_entreprises.empty and not df_mises.empty and not df_globale.empty:
    
    # Ici tu peux intégrer tout le calcul des KPIs comme dans les versions précédentes
    st.success("Tous les fichiers sont valides, les KPIs peuvent être calculés !")
    
    # Exemple rapide : total profils users
    st.metric("Total profils créés", len(df_users))
    
else:
    st.info("Veuillez uploader tous les fichiers correctement pour générer les KPIs.")
