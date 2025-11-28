import pandas as pd
from openpyxl import Workbook
from template_fiche_paie import generation_template_feuille_sans_rtt, generation_template_feuille_rtt

def fiche_paie(compte_travaux, regime_societe):
    compte_de_travaux = pd.read_excel(compte_travaux)
    employes = pd.read_excel(regime_societe)
    liste = []

    compte_de_travaux['Date'] = pd.to_datetime(compte_de_travaux['Date'])
    compte_de_travaux['Mois'] = compte_de_travaux['Date'].dt.month
    compte_de_travaux['Année'] = compte_de_travaux['Date'].dt.year
    mois = compte_de_travaux['Mois'].mode()[0]
    année = compte_de_travaux['Année'].mode()[0]

    heures_jour_personne = compte_de_travaux.groupby(['Nom', 'Prénom', 'Date'], as_index=False)['Heures'].sum()
    personnes = heures_jour_personne.groupby(['Nom', 'Prénom'])

    # Création de deux classeurs distincts
    wb_rtt = Workbook()
    ws0_rtt = wb_rtt.active
    ws0_rtt.title = "TEMP"

    wb_sans_rtt = Workbook()
    ws0_sans_rtt = wb_sans_rtt.active
    ws0_sans_rtt.title = "TEMP"

    for (nom, prenom), group in personnes:
        ligne = employes[
            (employes['Nom'].str.lower() == nom.lower()) &
            (employes['Prenom'].str.lower() == prenom.lower())
        ]
        if ligne.empty:
            liste.append((f"{nom} {prenom}"))
            continue

        entreprise = ligne['Entreprise'].values[0]
        regime = ligne['regime'].values[0]
        nom_feuille = f"{nom}_{prenom}"[:31]

        if regime == "rtt":
            ws = wb_rtt.create_sheet(title=nom_feuille)
            ws, date_line = generation_template_feuille_rtt(ws, nom, prenom, mois, année, entreprise)
        else:
            ws = wb_sans_rtt.create_sheet(title=nom_feuille)
            ws, date_line = generation_template_feuille_sans_rtt(ws, nom, prenom, mois, année, entreprise)

        for idx, ligne_group in group.iterrows():
            jour = ligne_group['Date']
            heures = ligne_group['Heures']
            if jour in date_line:
                ligne_excel = date_line[jour]
                ws[f'F{ligne_excel}'] = heures

    # Suppression des feuilles TEMP
    wb_rtt.remove(wb_rtt["TEMP"])
    wb_sans_rtt.remove(wb_sans_rtt["TEMP"])

    return wb_rtt, wb_sans_rtt, liste
    # Sauvegarde des deux fichiers séparés
    wb_rtt.save("10 - PAIE OCTOBRE - VALIDATION DES HEURES AVEC RTT 2025.xlsx")
    wb_sans_rtt.save("10 - PAIE OCTOBRE - VALIDATION DES HEURES SANS RTT 2025.xlsx")

    print("Fichiers Excel générés avec succès !")


if __name__ =="__main__":
    
    res = fiche_paie("../Paies/ExportCT OCTOBRE 2025-1.xlsx", "../regime_societe.xlsx")
