# -*- coding: utf8 -*-
import openpyxl
import json
"""
initialisation de variables
"""
#tout les membres du tableaux
listemembres = {}
#toutes les listes du tableaux des activitées
listelistes ={}
#toutes les activitées
listecards = {}
#tout les types d'activitées
listetypes = {}
#dico d'info à mettre dans l'excel pour une card
card = {"name":""
    ,"description":""
    ,"duedate":""
    ,"list":""
    ,"labels":""
    ,"membres":""}
#liste de dico card à parcourir pour l'import excel
cardsexcel = []
"""
ouverture du fichier json 
"""
with open("c:\\Users\\jazzt\\desktop\\NAM-IP\\Ih8ZoAQ4.json", "r") as file:
  trello = json.load(file)
listecards = trello["cards"]
listelistes= trello["lists"]
listemembres = trello["members"]
"""
recupération des listes à utilisées pour l'excel
dans une boucle
"""

for carte in listecards:
    #mettre les infos direct de la carte dans le dico
    card["name"] = carte["name"]
    card["description"] = carte["desc"]
    card["duedate"] = carte["due"]
    card["labels"] = carte["labels"]
    #chercher les id des différentes donnée externe à la carte
    for liste in listelistes:
        if carte["idList"] == liste["id"]:
            card["list"] = liste["name"]
    membrecard = []
    for membre in listemembres:
        for id in carte["idMembers"]:
             if id == membre["id"]:
                membrecard.append(membre["initials"])
    card["membres"] = membrecard
    print(card)
    #mise dans la liste pour l'excel
    cardsexcel.append(card)
"""
ouverture du fichier excel
"""
wb = openpyxl.load_workbook(filename = 'c:\\Users\\jazzt\\desktop\\NAM-IP\\tasks.xlsx')
ws = wb.create_sheet("tâches")
"""
titre des  colonnes + 
boucle pour importer les taches dans l'excel
"""
