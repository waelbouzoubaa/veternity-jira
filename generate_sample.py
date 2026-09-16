from openpyxl import Workbook

HEADERS = [
    "Ticket", "Lien Jira", "Titre", "Statut", "Date de création",
    "Date d'enrichissement", "Résumé RAG", "Ressources trouvées", "Score de pertinence",
]

rows = [
    [
        "IC-101", "https://veternity.atlassian.net/browse/IC-101",
        "Impossible de se connecter au module facturation", "En cours",
        "2026-09-14T09:12:00Z", "2026-09-14T09:15:00Z",
        "La documentation explique la procédure de réinitialisation du mot de passe "
        "et les prérequis réseau pour accéder au module.",
        "Procédure de connexion module facturation — https://confluence.veternity.com/pages/111\n"
        "FAQ Facturation — https://confluence.veternity.com/pages/112",
        0.82,
    ],
    [
        "IC-102", "https://veternity.atlassian.net/browse/IC-102",
        "Export PDF ne fonctionne pas sur Chrome", "À faire",
        "2026-09-14T10:03:00Z", "2026-09-14T10:06:00Z",
        "Aucun document pertinent n'a été trouvé pour ce sujet précis.",
        "Aucune ressource pertinente trouvée",
        0.31,
    ],
    [
        "IC-103", "https://veternity.atlassian.net/browse/IC-103",
        "Demande d'ajout d'un nouvel utilisateur", "Terminé",
        "2026-09-13T15:47:00Z", "2026-09-13T15:50:00Z",
        "La procédure de création d'utilisateur est documentée avec les rôles disponibles.",
        "Gestion des utilisateurs — https://confluence.veternity.com/pages/120",
        0.91,
    ],
    [
        "IC-104", "https://veternity.atlassian.net/browse/IC-104",
        "Erreur 500 lors de la synchro Salesforce", "En cours",
        "2026-09-15T08:20:00Z", "2026-09-15T08:23:00Z",
        "Quelques pistes de dépannage existent mais ne couvrent pas ce code d'erreur précis.",
        "Dépannage intégration Salesforce — https://confluence.veternity.com/pages/130",
        0.58,
    ],
]

wb = Workbook()
ws = wb.active
ws.title = "Tickets enrichis"
ws.append(HEADERS)
for row in rows:
    ws.append(row)

# Duplique le jeu de base avec des clés/dates différentes pour tester la pagination (>1 page)
for batch in range(1, 20):
    for row in rows:
        r = list(row)
        r[0] = f"{r[0]}-B{batch}"
        r[1] = r[1] + f"-B{batch}"
        ws.append(r)

wb.save("tickets_enrichis.xlsx")
print("tickets_enrichis.xlsx généré avec", ws.max_row - 1, "tickets de test")
