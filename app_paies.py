import streamlit as st
from paies import fiche_paie
import io

st.title("Export du compte de travaux en feuilles d'heures")

if "rtt_file" not in st.session_state:
    st.session_state.rtt_file = None

if "sans_rtt_file" not in st.session_state:
    st.session_state.sans_rtt_file = None

if "liste_manquants" not in st.session_state:
    st.session_state.liste_manquants = []

compte_travaux = st.file_uploader("Télécharger l'export du compte de travaux du mois", type="xlsx")
regime_societe = st.file_uploader("Télécharger l'excel des régimes et sociétés associés aux salariés", type="xlsx")

if st.button("Lancer le traitement"):
    if not compte_travaux or not regime_societe:
        st.error("Merci d'importer les deux fichiers avant de lancer le traitement.")
    else:
        try:
            wb_rtt, wb_sans_rtt, liste = fiche_paie(compte_travaux, regime_societe)

            buffer_rtt = io.BytesIO()
            wb_rtt.save(buffer_rtt)
            buffer_rtt.seek(0)

            buffer_sans_rtt = io.BytesIO()
            wb_sans_rtt.save(buffer_sans_rtt)
            buffer_sans_rtt.seek(0)

            st.session_state.rtt_file = buffer_rtt
            st.session_state.sans_rtt_file = buffer_sans_rtt
            st.session_state.liste_manquants = liste

            st.success("Les feuilles d'heures ont été générées avec succès.")

        except Exception as e:
            st.error(f"Une erreur a eu lieu : {e}")

if st.session_state.rtt_file:
    st.download_button(
        label="Télécharger feuilles avec RTT",
        data=st.session_state.rtt_file,
        file_name="feuilles_RTT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.session_state.sans_rtt_file:
    st.download_button(
        label="Télécharger feuilles sans RTT",
        data=st.session_state.sans_rtt_file,
        file_name="feuilles_sans_RTT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.session_state.liste_manquants:
    st.warning("Informations manquantes pour :")
    for nom in st.session_state.liste_manquants:
        st.write(f"- {nom}")
