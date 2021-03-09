# -*- coding: utf8 -*-
import configFichier
import os
import fonction
"""
initialisation de variables
"""
#tout les membres du tableaux
listemembres = {}
#toutes les listes du tableaux des activitées
listelistes ={}
#toutes les activitées
listecards = {}
"""
ouverture du fichier json 
"""
response = configFichier.requests.get("https://trello.com/b/Ih8ZoAQ4.json")
trello = configFichier.json.loads(response.text)
listecards = trello["cards"]
listelistes= trello["lists"]
listemembres = trello["members"]
"""
verification si le fichier existe déjà
"""
if os.path.isfile(configFichier.filename):
    print("update du fichier")
    wb = configFichier.openpyxl.load_workbook(configFichier.filename)
    ws1 = wb["tasks"]
    #backup de l'ancienne feuille de tasks
    ws1.title = "backuptasks"
    #création de la nouvelle feuille
    ws2 = wb.create_sheet(title="tasks")
    fonction.stylesheet(ws2)
    try:
        fonction.inserttasks(ws2, listecards, listemembres, listelistes)
            # wrap text pour la colonne des descriptions
        fonction.wrapcolumn(ws2, listecards)
        fonction.addfilter(ws2, len(listecards))
    finally:
        print("fin")
        wb.save(filename=configFichier.filename)
else:
 print("nouveau fichier")
 wb = configFichier.openpyxl.Workbook()
 ws = wb.active
 ws.title = "tasks"
 fonction.stylesheet(ws)
 try:
  fonction.inserttasks(ws, listecards, listemembres, listelistes)
   #wrap text pour la colonne des descriptions
  fonction.wrapcolumn(ws, listecards)
  fonction.addfilter(ws, len(listecards))
 finally:
     print("fin")
     wb.save(filename=configFichier.filename)
