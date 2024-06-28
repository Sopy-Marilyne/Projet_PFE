import streamlit as st
from utilities import extraire_titres_numerotes, add_data_to_existing_excel, inserer_sous_totaux
from io import BytesIO
import pandas as pd

def main():
    st.set_page_config(page_title="FAST DPGF", layout="wide")

    # Couleurs basées sur le logo
    primaryColor = "#8B4513"  # brun foncé
    secondaryColor = "#D4A76A"  # couleur dorée
    backgroundColor = "#FFF8DC"  # crème légère
    secondaryBackgroundColor = "#F5F5F5"  # gris très clair
    textColor = "#363636"
    font = "sans serif"

    st.markdown(
        f"""
        <style>
            @keyframes gradient {{
                0% {{ background-position: 0% 50%; }}
                50% {{ background-position: 100% 50%; }}
                100% {{ background-position: 0% 50%; }}
            }}
            .reportview-container {{
                background: linear-gradient(-45deg, {primaryColor}, {secondaryColor}, {backgroundColor}, {secondaryBackgroundColor});
                background-size: 400% 400%;
                animation: gradient 15s ease infinite;
                color: {textColor};
                font-family: {font};
            }}
            .sidebar .sidebar-content {{
                background: {secondaryBackgroundColor};
                padding: 20px;
                text-align: center;
            }}
            header .decoration {{
                background: {primaryColor};
            }}
            .stButton>button {{
                color: white;
                background-color: {primaryColor};
                border: none;
                padding: 10px 24px;
                border-radius: 8px;
                transition: background-color 0.3s ease;
            }}
            .stButton>button:hover {{
                background-color: {secondaryColor};
            }}
            .css-18e3th9 {{
                padding: 10px;
            }}
            .css-1d391kg {{
                padding-top: 3.5rem;
                padding-left: 1rem;
                padding-right: 1rem;
            }}
            .dataframe {{
                width: 100% !important;
                height: auto;
            }}
            .marquee {{
                font-size: 24px;
                color: white;
                background-color: {primaryColor};
                padding: 10px;
                border-radius: 5px;
                margin: 20px 0;
                width: 100%;
                overflow: hidden;
                position: relative;
            }}
            .marquee div {{
                display: inline-block;
                width: 100%;
                height: 100%;
                white-space: nowrap;
                animation: marquee 10s linear infinite;
            }}
            @keyframes marquee {{
                0%   {{ transform: translateX(100%); }}
                100% {{ transform: translateX(-100%); }}
            }}
            .sidebar-logo-container {{
                text-align: center;
            }}
            .sidebar-logo {{
                width: 80%;
                margin: 0 auto;
            }}
        </style>
        """,
        unsafe_allow_html=True
    )

    with st.sidebar:
        st.image("logo-cetab.jpg", width=200)
        st.markdown("""
            <p>
Le Groupe CETAB (Centre Etude Technique Aquitain du Bâtiment) est un Bureau d’études pluridisciplinaire spécialisé dans l’ingénierie du bâtiment, de l’infrastructure et de l’environnement.</p>
        """, unsafe_allow_html=True)

        st.write("## Paramètres")
        uploaded_file_word = st.file_uploader("Choisissez le contrat Word (.docx)", type=['docx'])
        uploaded_file_excel = st.file_uploader("Choisissez le template (.xlsx)", type=['xlsx'])

        # Champs de saisie pour les contenus des cellules A2 et C2
        cell_A2_content = st.text_input("NOM DU PROJET", "")
        cell_C2_content = st.text_input("N° du Lot", "xx - xxxxxxxxxxxxxxxxx")
        feuille_ = 'LOT ' + cell_C2_content[:2]
        feuille = "LOT XX"

    # Texte défilant
    st.markdown("""
    <div class="marquee">
        <div>Ceci est un outil interne au groupe CETAB permettant une extraction rapide des données.</div>
    </div>
    """, unsafe_allow_html=True)

    st.title("OUTIL DE DECOMPOSITION DE PRIX GLOBAL ET FORFAITAIRE")

    if uploaded_file_word:
        with st.spinner('Extraction des ouvrages...'):
            df_titres = extraire_titres_numerotes(uploaded_file_word)
            df_titres_ = df_titres.set_index('N°')
            
        # Ajout de boutons de suppression à chaque ligne
        if 'df_titres_' not in st.session_state:
            st.session_state.df_titres_ = df_titres_

        for i, row in st.session_state.df_titres_.iterrows():
            cols = st.columns((1, 4, 1))  # Ajustez la largeur des colonnes selon vos besoins
            cols[0].write(i)
            cols[1].write(row['DESIGNATION DES OUVRAGES'])
            if cols[2].button('Supprimer', key=f'delete_{i}'):
                st.session_state.df_titres_ = st.session_state.df_titres_.drop(i)

        st.dataframe(st.session_state.df_titres_, height=800, width=1200)

        if uploaded_file_excel and st.button('Télécharger le DPGF au format Excel'):
            # Convertir le fichier Excel téléchargé en BytesIO
            excel_data = BytesIO(uploaded_file_excel.read())
            df_modifie = inserer_sous_totaux(st.session_state.df_titres_)

            excel_output = add_data_to_existing_excel(df_modifie, cell_A2_content, cell_C2_content, feuille, feuille_, excel_data)
            st.download_button(label="📥 Télécharger",
                               data=excel_output,
                               file_name=f'DPGF_Lot_{cell_C2_content}.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == "__main__":
    main()
