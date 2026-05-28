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
