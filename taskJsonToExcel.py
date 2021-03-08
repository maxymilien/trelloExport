# -*- coding: utf8 -*-
import configFichier
import dateparser
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
with open(configFichier.filejson, "r",encoding="utf-8") as file:
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
    card['B'] = carte["name"]
    card['C'] = carte["desc"]
    d=dateparser.parse(str(carte["due"]))
    if d is None:
        card['D'] = ""
    else:
       card['D'] = str(d.day)+"/"+str(d.month)+"/"+str(d.year)
    for label in carte["labels"]:
           if lab == "":
               lab = str(label["name"])
           else:
               lab = str(lab)+","+str(label["name"])
    card['E'] =lab
    #chercher les id des différentes donnée externe à la carte
    for liste in listelistes:
        if carte["idList"] == liste["id"]:
            card['A'] = liste["name"]
    for membre in listemembres:
        for id in carte["idMembers"]:
             if id == membre["id"]:
               membreCard.append(membre["initials"])
    for item in membreCard:
        if mem == "":
            mem = str(item)
        else:
         mem = str(mem)+","+str(item)
    card['F'] = mem
    print(card)
    ws.column_dimensions['A'].width = 25
    ws["A1"] = "liste"
    ws.column_dimensions['B'].width = 65
    ws["B1"] = "nom tâche"
    ws.column_dimensions['C'].width = 140
    ws["C1"] = "description"
    ws.column_dimensions['D'].width = 25
    ws["D1"] = "echéance"
    ws.column_dimensions['E'].width = 35
    ws["E1"] = "étiquettes"
    ws["F1"] = "équipe"
    ws.append(card)
    card.clear()
finally:
    print("fin")
    wb.save(filename=configFichier.filexlsx)
