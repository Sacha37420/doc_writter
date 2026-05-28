"""
Exemple complet d'utilisation de doc_writer.

Génère un rapport en Word (.docx) et en HTML à partir du même contenu,
en utilisant le template de page "standard_report_page" et en appelant
directement des templates combinés.
"""

import sys
from pathlib import Path

# Permet d'exécuter l'exemple sans installer le package
sys.path.insert(0, str(Path(__file__).parent.parent))

from doc_writer import HtmlDocumentWriter, WordDocumentWriter

# ---------------------------------------------------------------------------
# Contenu partagé Word + HTML
# ---------------------------------------------------------------------------

SECTIONS = [
    # --- hero_section : titre + paragraphe (60 %) + image (40 %) ---
    {
        "title": {"text": "Rapport Annuel 2024", "level": 1},
        "paragraph": {
            "text": (
                "Ce rapport présente les résultats et les indicateurs clés de "
                "l'exercice 2024, avec une analyse détaillée des performances "
                "par département et les perspectives pour 2025."
            )
        },
        "image": {"caption": "Figure 1 : Vue d'ensemble du groupe", "path": None},
    },
    # --- content_with_list : titre + paragraphe + liste à puces ---
    {
        "title": {"text": "Points Clés de l'Exercice", "level": 2},
        "paragraph": {
            "text": "L'exercice 2024 a été marqué par plusieurs avancées significatives :"
        },
        "bullet_list": {
            "items": [
                "Croissance du chiffre d'affaires de 15 % par rapport à 2023",
                "Lancement de 3 nouveaux produits sur le marché européen",
                "Expansion dans 2 nouveaux pays : Belgique et Portugal",
                "Réduction de l'empreinte carbone de 20 % grâce aux initiatives vertes",
                "Taux de satisfaction client en hausse à 94 %",
            ]
        },
    },
    # --- data_table_section : titre + tableau ---
    {
        "title": {"text": "Indicateurs Financiers", "level": 2},
        "table": {
            "headers": ["Indicateur", "2023", "2024", "Variation"],
            "rows": [
                ["Chiffre d'affaires", "2,4 M€", "2,76 M€", "+15 %"],
                ["Résultat net", "380 K€", "456 K€", "+20 %"],
                ["Marge brute", "58 %", "61 %", "+3 pts"],
                ["Effectif total", "42", "51", "+9"],
            ],
        },
    },
]

# ---------------------------------------------------------------------------
# 1. Génération Word via template de page
# ---------------------------------------------------------------------------

word = WordDocumentWriter()
word.add_page("standard_report_page", SECTIONS)
word.save("output/rapport_2024.docx")
print("✓ Word  →  examples/output/rapport_2024.docx")

# ---------------------------------------------------------------------------
# 2. Génération HTML via template de page
# ---------------------------------------------------------------------------

html = HtmlDocumentWriter()
html.add_page("standard_report_page", SECTIONS)
html.save("output/rapport_2024.html")
print("✓ HTML  →  examples/output/rapport_2024.html")

# ---------------------------------------------------------------------------
# 3. Utilisation directe de templates combinés (sans page template)
# ---------------------------------------------------------------------------

word2 = WordDocumentWriter()
word2.add_combined(
    "cover_section",
    {
        "image": {"caption": "Logo de l'entreprise", "path": None},
        "title": {"text": "Guide Utilisateur v2.0", "level": 1},
        "paragraph": {"text": "Document confidentiel — usage interne uniquement."},
    },
)
word2.add_combined(
    "content_with_list",
    {
        "title": {"text": "Table des matières", "level": 2},
        "paragraph": {"text": "Ce guide couvre les rubriques suivantes :"},
        "bullet_list": {
            "items": ["Installation", "Configuration", "Utilisation avancée", "FAQ"]
        },
    },
)
word2.save("output/guide_utilisateur.docx")
print("✓ Word  →  examples/output/guide_utilisateur.docx")

html2 = HtmlDocumentWriter()
html2.add_combined(
    "cover_section",
    {
        "image": {"caption": "Logo de l'entreprise", "path": None},
        "title": {"text": "Guide Utilisateur v2.0", "level": 1},
        "paragraph": {"text": "Document confidentiel — usage interne uniquement."},
    },
)
html2.add_combined(
    "content_with_list",
    {
        "title": {"text": "Table des matières", "level": 2},
        "paragraph": {"text": "Ce guide couvre les rubriques suivantes :"},
        "bullet_list": {
            "items": ["Installation", "Configuration", "Utilisation avancée", "FAQ"]
        },
    },
)
html2.save("output/guide_utilisateur.html")
print("✓ HTML  →  examples/output/guide_utilisateur.html")

# ---------------------------------------------------------------------------
# 4. Graphique — chart_section
# ---------------------------------------------------------------------------

word3 = WordDocumentWriter()
word3.add_combined(
    "chart_section",
    {
        "title": {"text": "Évolution du chiffre d'affaires", "level": 2},
        "chart": {
            "chart_type": "bar",
            "labels": ["T1", "T2", "T3", "T4"],
            "datasets": [
                {"label": "2023", "values": [540, 620, 580, 700]},
                {"label": "2024", "values": [610, 695, 670, 810]},
            ],
        },
    },
)
word3.save("output/graphique.docx")
print("✓ Word  →  examples/output/graphique.docx")

html3 = HtmlDocumentWriter()
html3.add_combined(
    "chart_section",
    {
        "title": {"text": "Évolution du chiffre d'affaires", "level": 2},
        "chart": {
            "chart_type": "bar",
            "labels": ["T1", "T2", "T3", "T4"],
            "datasets": [
                {"label": "2023", "values": [540, 620, 580, 700]},
                {"label": "2024", "values": [610, 695, 670, 810]},
            ],
        },
    },
)
html3.save("output/graphique.html")
print("✓ HTML  →  examples/output/graphique.html")

# ---------------------------------------------------------------------------
# 5. Logigramme simple (formes Word natives — recommandé)
# ---------------------------------------------------------------------------

FLOWCHART_CONTENT = {
    "title": {"text": "Processus de traitement des demandes", "level": 2},
    "flowchart": {
        "flowchart_type": "simple",   # ← toujours préférer simple
        "nodes": [
            {"id": "1", "type": "start",    "text": "Réception de la demande"},
            {"id": "2", "type": "process",  "text": "Vérifier les informations"},
            {"id": "3", "type": "decision", "text": "Dossier complet ?"},
            {"id": "4", "type": "io",       "text": "Enregistrer dans le système"},
            {"id": "5", "type": "process",  "text": "Notifier le demandeur"},
            {"id": "6", "type": "end",      "text": "Traitement terminé"},
        ],
    },
}

word4 = WordDocumentWriter()
word4.add_combined("flowchart_section", FLOWCHART_CONTENT)
word4.save("output/logigramme.docx")
print("✓ Word  →  examples/output/logigramme.docx")

html4 = HtmlDocumentWriter()
html4.add_combined("flowchart_section", FLOWCHART_CONTENT)
html4.save("output/logigramme.html")
print("✓ HTML  →  examples/output/logigramme.html")
