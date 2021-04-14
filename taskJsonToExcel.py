# -*- coding: utf8 -*-
import configuration
import os
import utilitaire
"""
initialisation de variables
"""
#choix couleur ou noir et blanc
couleur = ""
#tout les membres du tableaux
listemembres = {}
#toutes les listes du tableaux des activitées
listelistes ={}
#toutes les activitées
listecards = {}
"""
ouverture du fichier json 
"""
#with open(configuration.file, "r", encoding="utf-8") as file:
 #trello = configuration.json.load(file)
response = configuration.requests.get(url=configuration.url)
trello = configuration.json.loads(response.text)
listecards = trello["cards"]
listelistes= trello["lists"]
listemembres = trello["members"]
"""
verification si le fichier existe déjà
"""
if os.path.isfile(configuration.filename):
    print("update du fichier")
    wb = configuration.openpyxl.load_workbook(configuration.filename)
    ws1 = wb["tasks"]
    #backup de l'ancienne feuille de tasks
    ws1.title = "backuptasks"
    #création de la nouvelle feuille
    ws2 = wb.create_sheet(title="tasks")
    utilitaire.stylesheet(ws2)
    couleur = input('entre (n) pour la version noir et blanc et (c) pour la version couleur')
    try:
        utilitaire.inserttasks(ws2, listecards, listemembres, listelistes)
            # wrap text pour la colonne des descriptions
        utilitaire.wrapcolumn(ws2, len(listecards))
        #changement du format de la  dat
        utilitaire.dateColumn(ws2, len(listecards))
        if str(couleur).lower() == "c":
          #ajout des rêgles couleur
            utilitaire.addrulesCouleur(ws2, len(listecards))
        else:
            # ajout des rêgles noir et blanc
            utilitaire.addrulesBW(ws2)
    finally:
        print("fin")
        wb.save(filename=configuration.filename)
else:
    print("nouveau fichier")
    wb = configuration.openpyxl.Workbook()
    ws = wb.active
    ws.title = "tasks"
    utilitaire.stylesheet(ws)
    couleur = input('entre (n) pour la version noir et blanc et (c) pour la version couleur')
    try:
        utilitaire.inserttasks(ws, listecards, listemembres, listelistes)
        #wrap text pour la colonne des descriptions
        utilitaire.wrapcolumn(ws, len(listecards))
        # changement du format de la date
        utilitaire.dateColumn(ws, len(listecards))
        if str(couleur).lower() == "c":
          #ajout des rêgles couleur
            utilitaire.addrulesCouleur(ws, len(listecards))
        else:
            # ajout des rêgles noir et blanc
            utilitaire.addrulesBW(ws)
    finally:
        print("fin")
        wb.save(filename=configuration.filename)
