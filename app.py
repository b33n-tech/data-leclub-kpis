import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.title("Dashboard Marketplace & Incubateur")

# --- Upload des fichiers ---
st.sidebar.header("Uploader les fichiers")
file_users = st.sidebar.file_uploader("Noms persos.csv", type=["csv"])
file_entreprises = st.sidebar.file_uploader("Entreprises données.xlsx", type=["xlsx"])
file_mises_relation = st.sidebar.file_uploader("Historique des mises en relation.csv", type=["csv"])
file_base_globale = st.sidebar.file_uploader("Base globale projet.xlsx", type=["xlsx"])

if file_users and file_entreprises and file_mises_relation and file_base_globale:
    # --- Lecture des fichiers ---
    df_users = pd.read_csv(file_users)
    df_entreprises = pd.read_excel(file_entreprises)
    df_mises = pd.read_csv(file_mises_relation)
    df_globale = pd.read_excel(file_base_globale)

    # --- Datas globales ---
    st.header("Datas globales")
    demandes_total = len(df_mises)
    profils_total = len(df_users)
    
    # Profils connectés sur le mois courant
    today = datetime.today()
    month_ago = today - timedelta(days=30)
    df_users['Date de dernière connexion'] = pd.to_datetime(df_users['Date de dernière connexion'], dayfirst=True, errors='coerce')
    profils_connectes = df_users[df_users['Date de dernière connexion'] >= month_ago].shape[0]
    
    st.metric("Demandes de mise en relation", demandes_total)
    st.metric("Profils créés", profils_total)
    st.metric("Profils connectés sur le mois", profils_connectes)

    # --- Marketplace ---
    st.header("Marketplace")
    go_between_valides = (df_mises['Go between validé'] == 'Oui').sum()
    rdv_realises = df_mises['RDV réalisés'].sum()
    rdv_non_realises = df_mises['Rdv non réalisé'].sum()
    
    st.subheader("Totaux")
    st.metric("Go Between validés", go_between_valides)
    st.metric("RDV réalisés", rdv_realises)
    st.metric("RDV non réalisés", rdv_non_realises)

    # Statut des mises en relation
    st.subheader("Statut des mises en relation")
    statuts_counts = df_mises['Statut des mises en relation à date'].value_counts()
    st.table(statuts_counts)

    # Taux de conversion (moyenne)
    taux_go_between = pd.to_numeric(df_mises['Taux de conversion goBetween'], errors='coerce').mean()
    taux_rdv = pd.to_numeric(df_mises['Taux de conversion RDV réalisé'], errors='coerce').mean()
    st.metric("Taux de conversion Go Between (%)", round(taux_go_between, 2))
    st.metric("Taux de conversion RDV réalisés (%)", round(taux_rdv, 2))

    # Répartition trimestrielle des demandes
    df_mises['Dates simples'] = pd.to_datetime(df_mises['Dates simples'], format='%Y-%m', errors='coerce')
    df_mises['Trimestre'] = df_mises['Dates simples'].dt.to_period('Q')
    trimestriel = df_mises.groupby('Trimestre').size()
    st.subheader("Répartition trimestrielle des demandes")
    st.table(trimestriel)

    # --- Profils persos & Sociétés ---
    st.header("Profils persos & Sociétés")
    st.metric("Nombre total d'entrepreneurs", len(df_globale))
    st.metric("Total profils persos", len(df_users))

    st.subheader("Statut profils persos")
    st.table(df_users['Statut'].value_counts())

    st.subheader("Statut profils sociétés")
    st.table(df_entreprises['Statut'].value_counts())

    # --- Complétion des profils (Base Globale) ---
    st.header("Complétion des profils")
    st.subheader("Vue globale")
    st.table(df_globale['Profil personnel Le Club'].value_counts())
    st.table(df_globale['Profil sociétés Le Club'].value_counts())

    st.subheader("Par CAR/SUM")
    car_group = df_globale.groupby('CAR/SUM (territorial)')['Profil sociétés Le Club'].value_counts()
    st.table(car_group)

    st.subheader("Par Incubateur territorial")
    incub_group = df_globale.groupby('Incubateur territorial')['Profil sociétés Le Club'].value_counts()
    st.table(incub_group)

    st.subheader("% de complétion sur les profils incubation individuelle")
    incubation_indiv = df_globale[df_globale['Statut d'incubation'] == 'Incubation individuelle']
    st.table(incubation_indiv['Profil sociétés Le Club'].value_counts(normalize=True).mul(100).round(2))
    
else:
    st.info("Veuillez uploader les 4 fichiers pour générer les KPIs.")
