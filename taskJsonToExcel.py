# -*- coding: utf8 -*-
import configFichier
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
card = {}
"""
ouverture du fichier json 
"""
with open(configFichier.filejson, "r") as file:
  trello = configFichier.json.load(file)
listecards = trello["cards"]
listelistes= trello["lists"]
listemembres = trello["members"]
"""
recupération des listes à utilisées pour l'excel
dans une boucle
"""
wb = configFichier.openpyxl.load_workbook(configFichier.filexlsx)
ws = wb['tasks']
try:
 for carte in listecards:
    membreCard = []
    lab = ""
    mem = ""
    #mettre les infos direct de la carte dans le dico
    card["A"] = carte["name"]
    card["E"] = carte["desc"]
    card["M"] = carte["due"]
    for label in carte["labels"]:
           lab = str(lab)+","+str(label["name"])
    card["O"] =lab
    #chercher les id des différentes donnée externe à la carte
    for liste in listelistes:
        if carte["idList"] == liste["id"]:
            card["S"] = liste["name"]
    for membre in listemembres:
        for id in carte["idMembers"]:
             if id == membre["id"]:
               membreCard.append(membre["initials"])
    for item in membreCard:
        mem = str(mem)+","+str(item)
    card['V'] = mem
    print(card)
    ws.append(card)
    card.clear()
finally:
 wb.save(filename=configFichier.filexlsx)