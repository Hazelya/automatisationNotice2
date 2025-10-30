import shutil
import zipfile
from pathlib import Path
import streamlit as st
import pandas as pd
import os
import datetime as dt
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML



# === Fonctions utiles ===
def date_now():
    mois = ['', 'janvier', 'février', 'mars', 'avril', 'mai', 'juin',
            'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
    today = dt.date.today()
    return f"{today.day} {mois[today.month]} {today.year}"

def format_nombre(nombre):
    return f"{nombre:,.2f}".replace(',', ' ').replace('.', ',')


# === Interface Streamlit ===
st.title("Générateur de notices d'appel de fonds")

uploaded_file = st.file_uploader("Fichier Excel de données", type=["xlsx"])
header = st.text_input("Première ligne (header)", value="3")
header = int(header) - 1
# texte_fond_couvrir = st.text_area("Texte pour couvrir l'appel")
texte_fond_finance = st.text_area("Texte")

numero_call = st.text_input("Numéro de l'appel", value="9")

date = st.text_input("Date d'envoie", value="30/11/2025")
date_obj = dt.datetime.strptime(date, "%d/%m/%Y")

date_call = st.text_input("Date du CALL", value="17/11/2025")
pourcentage_call = st.text_input("Pourcentage du CALL", value="10,50")
pourcentage_avant_call = st.text_input("Pourcentage du pré CALL", value="87,00")
nom_fond = st.text_input("Nom du fonds", value="FPCI ÉPOPÉE Xplore II")
pays = st.text_input("Pays", value="France")

if st.button("Générer les notices"):
    if uploaded_file:
        # Sauvegarde temporaire
        os.makedirs("ressources", exist_ok=True)
        chemin_fichier = "ressources/Base-data-test-fund-exercice.xlsx"
        with open(chemin_fichier, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            df = pd.read_excel(chemin_fichier, sheet_name='SOUSCRIPTEURS', header=header) # Header a modifier si besoin
            df_nettoye = df[df['SOUSCRIPTEUR'].notna()]
            df_nettoye = df_nettoye[~df_nettoye['SOUSCRIPTEUR'].str.startswith('TOTAL', na=False)]
            df_nettoye = df_nettoye.reset_index(drop=True)

            df_temp = pd.read_excel(chemin_fichier, header=None)

            # Récupérer les deux premières cellules
            iban = df_temp.iloc[0, 1]  # B1
            bic = df_temp.iloc[1, 1]  # B2

            # df_CALL = pd.read_excel(chemin_fichier, sheet_name='SOUSCRIPTEURS', header=3) # ne sert plus
            call = 'CALL #' + numero_call
            montant_total = df[call][df.shape[0]-6]
            #date_call = df_CALL.loc[df_CALL['Nominal'] == call, 'Date'].iloc[0]
            #pourcentage_call = df_CALL.loc[df_CALL['Nominal'] == call, df_CALL.columns[2]].iloc[0]

            dir = 'ressources'
            env = Environment(loader=FileSystemLoader(dir))
            template = env.get_template('model_notice_img.html')

            for folder in ["Output", "Output_HTML"]:
                if os.path.exists(folder):
                    shutil.rmtree(folder)

            os.makedirs('Output', exist_ok=True)
            os.makedirs('Output_HTML', exist_ok=True)

            for i in range(df_nettoye.shape[0]):
                # total_avant_call = df_nettoye['TOTAL APPELE'][i] - df_nettoye[call][i]
                # pourcentage_avant_call = (total_avant_call / df_nettoye['ENGAGEMENT'][i]) * 100

                if pd.isna(df_nettoye["Représentant"][i]):
                    representant = ''
                else:
                    representant = df_nettoye["Représentant"][i]

                # Les données à injecter
                # 'balise' : 'la donnée',
                data = {
                    'souscripteur': df_nettoye["SOUSCRIPTEUR"][i],
                    # 'pm_pp': df_nettoye["TYPE"][i],
                    'representant': representant,
                    'adresse': df_nettoye["ADRESSE"][i],
                    'code_postal': str(df_nettoye["CP"][i]),
                    'ville': df_nettoye["VILLE"][i],
                    'pays': pays,
                    'date': str(date_obj.strftime("%d %B %Y")),
                    'numero_call': numero_call,
                    'date_call': date_call,
                    'nom_fond': nom_fond,
                    'montant_total': format_nombre(montant_total),
                    # 'pourcentage_call': f"{pourcentage_call * 100:.2f}",
                    'pourcentage_call': pourcentage_call,
                    'montant_a_liberer': format_nombre(df_nettoye[call][i]),
                    'pourcentage_avant_call': pourcentage_avant_call,
                    # 'texte_fond_couvrir': texte_fond_couvrir,
                    'texte_fond_finance': texte_fond_finance,
                    'montant_engagement_initial': format_nombre(df_nettoye["ENGAGEMENT"][i]),
                    'nombre_parts_souscrites': format_nombre(df_nettoye["NBR PARTS"][i]),
                    'categorie_part': df_nettoye["PART"][i],
                    'total_appele': format_nombre(df_nettoye["TOTAL APPELE"][i]),
                    'pourcent_liberation': f"{df_nettoye['%LIBERATION'][i] * 100:.2f}",
                    'residuel': format_nombre(df_nettoye["RESIDUEL"][i]),
                    'libelle_virement': 'CR ' + df_nettoye["SOUSCRIPTEUR"][i] + ' ADF ' + numero_call,
                    'iban': iban,
                    'bic': bic,
                }

                # Rend le HTML final avec tes vraies données
                html_content = template.render(data)


                date_call_obj = dt.datetime.strptime(date_call, "%d/%m/%Y")
                date_title = date_obj.strftime("%Y%m%d")


                # Sauve le résultat dans un fichier
                os.makedirs('Output_HTML', exist_ok=True)
                dir_nom_fichier = 'Output_HTML/' + str(date_title) + '_' + df_nettoye["SOUSCRIPTEUR"][i] + '.html'
                with open(dir_nom_fichier, 'w', encoding='utf-8') as file:
                    file.write(html_content)

                # print('Notice HTML générée avec succès.')

                os.makedirs('Output/PDF', exist_ok=True)
                os.makedirs('Output/Word', exist_ok=True)

                fichier_html = 'Output_HTML/' + str(date_title) + '_' + df_nettoye["SOUSCRIPTEUR"][i] + '.html'
                fichier_pdf = 'Output/PDF/' + str(date_title) + '_' + df_nettoye["SOUSCRIPTEUR"][i] + '.pdf'

                base_url = Path('ressources/images').resolve()  # Chemin absolu vers /ressources

                HTML(filename=fichier_html, base_url=base_url.as_uri()).write_pdf(fichier_pdf)


            # Zip tous les fichiers
            shutil.make_archive("notices", "zip", "Output")
            with open("notices.zip", "rb") as f:
                st.download_button("Télécharger les notices générées", f, "notices.zip")

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
    else:
        st.warning("Merci de déposer un fichier Excel avant de lancer la génération.")
