import configFichier

def stylesheet(ws):
    # faire un wrap sur description et nom tache
    # colonne des listes
    ws.column_dimensions['A'].width = 25
    ws["A1"] = "Liste"
    ws["A1"].font = configFichier.font
    # colonne des tâches
    ws.column_dimensions['B'].width = 65
    ws["B1"] = "Nom tâche"
    ws["B1"].font = configFichier.font
    # colonne des descriptions
    ws.column_dimensions['C'].width = 60
    ws["C1"] = "Description"
    ws["C1"].font = configFichier.font
    # colonne des échéances
    ws.column_dimensions['D'].width = 25
    ws["D1"] = "Echéance"
    ws["D1"].font = configFichier.font
    # colonne des étiquettes
    ws.column_dimensions['E'].width = 35
    ws["E1"] = "Etiquettes"
    ws["E1"].font = configFichier.font
    # colonne des équipes
    ws["F1"] = "Equipe"
    ws["F1"].font = configFichier.font

def wrapcolumn(ws, listecards):
    for row in ws.iter_rows(min_row=1, max_row=len(listecards), min_col=3):
        for cell in row:
            cell.alignment = configFichier.wrap

def inserttasks(ws, listecards, listemembres, listelistes):
    # dico d'info à mettre dans l'excel pour une card
    card = {}
    for carte in listecards:
        membreCard = []
        lab = ""
        mem = ""
        # mettre les infos direct de la carte

        # recuperation du nom de la tâche
        card['B'] = carte["name"]

        # récuperation de la description
        card['C'] = carte["desc"]

        # récuperation de la date d'échéance + formatage
        d = configFichier.dateparser.parse(str(carte["due"]))
        if d is None:
            card['D'] = ""
        else:
            card['D'] = str(d.day) + "-" + str(d.month) + "-" + str(d.year)

        # récuperation des étiquettes de la tâche + formatage en string
        for label in carte["labels"]:
            if lab == "":
                lab = str(label["name"])
            else:
                lab = str(lab) + "," + str(label["name"])
        card['E'] = lab

        # chercher les id des différentes donnée externe à la carte

        # récupération de la liste de la tâche
        for liste in listelistes:
            if carte["idList"] == liste["id"]:
                card['A'] = liste["name"]

        # récuperation des membres assignés
        for membre in listemembres:
            for id in carte["idMembers"]:
                if id == membre["id"]:
                    membreCard.append(membre["initials"])

        # formatage en string de la liste des membres
        for item in membreCard:
            if mem == "":
                mem = str(item)
            else:
                mem = str(mem) + "," + str(item)
        card['F'] = mem
        ws.append(card)
        card.clear()

def addfilter(ws,longueur):
    ws.auto_filter.ref = "A1:F"+str(longueur)
    ws.auto_filter.add_sort_condition("A2:F"+str(longueur))