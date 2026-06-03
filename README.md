# doc_writer

Module Python pour générer des documentations **Word (.docx)** et **HTML** à partir de templates JSON personnalisés.

---

## Concepts

Le module repose sur trois niveaux de templates imbriqués :

```
Templates singuliers  ──▶  Templates combinés  ──▶  Templates de pages
(blocs atomiques)          (mise en page)             (page A4 complète)
```

### 1. Templates singuliers

Blocs atomiques dont le style (police, taille, espacement) est **fixé une fois pour toutes** dans le JSON. Seules les **couleurs** sont paramétrables via des placeholders `{{nom_couleur}}` résolus à partir de la palette du template combiné parent.

| Template | Fichier | Description |
|----------|---------|-------------|
| Paragraphe | `singular/paragraph.json` | Bloc de texte courant, justifié |
| Liste à puces | `singular/bullet_list.json` | Liste avec puce et retrait fixes, justifiée |
| Titre | `singular/title.json` | Niveaux 1, 2 et 3 |
| Image | `singular/image.json` | Cadre fixe avec placeholder |
| Tableau | `singular/table.json` | Nb colonnes/lignes fixés, style de bordure fixé |
| Encadré | `singular/callout.json` | Bloc mis en valeur avec bordure gauche colorée |
| Séparateur | `singular/separator.json` | Ligne horizontale de séparation |
| Graphique | `singular/chart.json` | Graphique matplotlib (bar, barh, line, pie) |
| Logigramme | `singular/flowchart.json` | Logigramme simple (Word natif) ou complexe (SVG/PNG) |

#### Placeholders de couleur disponibles

| Placeholder | Rôle |
|-------------|------|
| `{{title_color}}` | Couleur des titres |
| `{{text_color}}` | Couleur du texte courant |
| `{{bullet_color}}` | Couleur de la puce |
| `{{list_text_color}}` | Couleur du texte des items de liste |
| `{{image_border_color}}` | Couleur du cadre image |
| `{{table_header_bg}}` | Couleur de fond de l'en-tête de tableau |
| `{{table_border_color}}` | Couleur des bordures du tableau |
| `{{callout_border_color}}` | Couleur de la bordure gauche de l'encadré |
| `{{callout_bg_color}}` | Couleur de fond de l'encadré |
| `{{chart_color_1}}` … `{{chart_color_4}}` | Couleurs des séries de graphiques |

---

### 2. Templates combinés

Combinent **2 à 5 templates singuliers** dans une grille de lignes/colonnes.  
Chaque template combiné définit :
- une **palette de couleurs** qui résout tous les `{{placeholders}}` des singuliers
- un **layout en grille** (`rows` → `columns` → `components`)

#### Templates combinés fournis

| ID | Mise en page | Palette |
|----|-------------|---------|
| `hero_section` | Titre (100 %) / Paragraphe (60 %) + Image (40 %) | Bleue |
| `content_with_list` | Titre + Paragraphe + Liste à puces (100 %) | Verte |
| `data_table_section` | Titre + Tableau (100 %) | Violette |
| `cover_section` | Image + Grand titre + Sous-titre (100 %) | Marine/or |
| `two_column_text` | Sous-titre + Paragraphe × 2 colonnes 50/50 | Teal |
| `chart_section` | Titre + Graphique (100 %) | Bleue |
| `flowchart_section` | Titre + Logigramme (100 %) | Bleue |

---

### 3. Templates de pages

Combinent plusieurs templates combinés pour former une **page A4 portrait** complète.

| ID | Sections dans l'ordre |
|----|-----------------------|
| `standard_report_page` | `hero_section` → `content_with_list` → `data_table_section` |
| `cover_page` | `cover_section` |

---

## Installation

```bash
pip install -e .
```

Dépendances : **python-docx ≥ 1.1.0** · **matplotlib ≥ 3.7** (graphiques et logigrammes complexes)

> `cairosvg` est optionnel : s'il est installé, les logigrammes complexes sont convertis en PNG haute fidélité depuis le SVG. Sans lui, le rendu Word utilise matplotlib comme fallback.

---

## Utilisation Python

### Via JSON de document (recommandé)

```python
from doc_writer import WordDocumentWriter, HtmlDocumentWriter

word = WordDocumentWriter()
word.write_from_json("mon_document.json")   # chemin vers un fichier JSON
word.save("sortie.docx")

html = HtmlDocumentWriter()
html.write_from_json("mon_document.json")
html.save("sortie.html")
```

`write_from_json` accepte aussi un dict déjà parsé ou une chaîne JSON brute.

### Via un template de page

```python
word = WordDocumentWriter()
word.add_page("standard_report_page", [
    {  # section 1 — hero_section
        "title":     {"text": "Rapport Annuel 2024", "level": 1},
        "paragraph": {"text": "Résultats de l'exercice…"},
        "image":     {"caption": "Figure 1", "path": None},
    },
    {  # section 2 — content_with_list
        "title":       {"text": "Points Clés", "level": 2},
        "paragraph":   {"text": "Les faits marquants :"},
        "bullet_list": {"items": ["Croissance de 15 %", "3 nouveaux produits"]},
    },
    {  # section 3 — data_table_section
        "title": {"text": "Indicateurs", "level": 2},
        "table": {
            "headers": ["Indicateur", "2023", "2024"],
            "rows": [["CA", "2,4 M€", "2,76 M€"]],
        },
    },
])
word.save("rapport.docx")
```

### Via des templates combinés directement

```python
word = WordDocumentWriter()
word.add_combined("hero_section", {
    "title":     {"text": "Introduction", "level": 1},
    "paragraph": {"text": "Texte principal."},
    "image":     {"caption": "Illustration", "path": None},
})
word.add_combined("two_column_text", {
    "title":     [{"text": "Avantages", "level": 3}, {"text": "Inconvénients", "level": 3}],
    "paragraph": [{"text": "Texte gauche."}, {"text": "Texte droit."}],
})
word.save("document.docx")
```

### Avec templates personnalisés

```python
word = WordDocumentWriter(custom_templates_dir="/chemin/vers/mes_templates")
```

---

## Format JSON de document

> Cette section est la **référence normative** du format attendu par `write_from_json`.  
> Elle est rédigée de façon à permettre à une IA de générer des JSONs valides sans consulter le code.

### Structure racine

```json
{
  "title": "string",
  "blocks": [ ... ]
}
```

| Champ | Type | Requis | Description |
|-------|------|--------|-------------|
| `title` | string | non | Métadonnée descriptive, non affichée dans le document |
| `blocks` | tableau | **oui** | Liste ordonnée des blocs à générer, du premier au dernier |

---

### Bloc de type `"page"`

Rend un template de page complet (plusieurs sections d'un coup).

```json
{
  "type": "page",
  "template": "<id_page>",
  "sections": [ <contenu_section_0>, <contenu_section_1>, ... ]
}
```

| Champ | Type | Requis | Description |
|-------|------|--------|-------------|
| `type` | `"page"` | **oui** | Discriminant de type |
| `template` | string | **oui** | ID du template de page |
| `sections` | tableau | **oui** | Un objet de contenu par section, dans l'ordre du template |

**Templates de page disponibles**

| `template` | Nombre de sections | Sections dans l'ordre |
|---|---|---|
| `"standard_report_page"` | 3 | `hero_section`, `content_with_list`, `data_table_section` |
| `"cover_page"` | 1 | `cover_section` |

---

### Bloc de type `"combined"`

Rend un seul template combiné.

```json
{
  "type": "combined",
  "template": "<id_combiné>",
  "content": { ... }
}
```

| Champ | Type | Requis | Description |
|-------|------|--------|-------------|
| `type` | `"combined"` | **oui** | Discriminant de type |
| `template` | string | **oui** | ID du template combiné |
| `content` | objet | **oui** | Contenu des composants (voir référence par template) |

---

### Référence des templates combinés

Pour chaque template, la table indique les **clés exactes** attendues dans `content`, leur multiplicité (`Nb`) et le `level` par défaut des titres.

> **Règle des multiples :** si `Nb > 1`, la valeur DOIT être un tableau de `Nb` objets dans l'ordre gauche→droite, haut→bas. Si `Nb = 1`, la valeur peut être un objet seul `{}` ou un tableau à un élément `[{}]`.

#### `hero_section` — palette bleue

Mise en page : **Titre** pleine largeur / **Paragraphe** (60 %) + **Image** (40 %)

| Clé `content` | Nb | `level` par défaut | Description |
|---|---|---|---|
| `title` | 1 | 1 | Titre de la section |
| `paragraph` | 1 | — | Texte principal |
| `image` | 1 | — | Image ou placeholder |

#### `content_with_list` — palette verte

Mise en page : **Titre** / **Paragraphe** / **Liste à puces** (tout pleine largeur)

| Clé `content` | Nb | `level` par défaut | Description |
|---|---|---|---|
| `title` | 1 | 2 | Titre de la section |
| `paragraph` | 1 | — | Texte d'introduction |
| `bullet_list` | 1 | — | Liste à puces |

#### `data_table_section` — palette violette

Mise en page : **Titre** / **Tableau** (tout pleine largeur)

| Clé `content` | Nb | `level` par défaut | Description |
|---|---|---|---|
| `title` | 1 | 2 | Titre de la section |
| `table` | 1 | — | Tableau de données |

#### `cover_section` — palette marine/or

Mise en page : **Image** / **Titre** / **Paragraphe** (tout pleine largeur)

| Clé `content` | Nb | `level` par défaut | Description |
|---|---|---|---|
| `image` | 1 | — | Visuel de couverture ou placeholder |
| `title` | 1 | 1 | Titre principal du document |
| `paragraph` | 1 | — | Sous-titre ou description |

#### `two_column_text` — palette teal

Mise en page : deux colonnes 50/50, chacune avec **Sous-titre** + **Paragraphe**

| Clé `content` | Nb | `level` par défaut | Description |
|---|---|---|---|
| `title` | **2** | 3 | `[0]` = colonne gauche · `[1]` = colonne droite |
| `paragraph` | **2** | — | `[0]` = texte gauche · `[1]` = texte droit |

#### `chart_section` — palette bleue

Mise en page : **Titre** / **Graphique** (tout pleine largeur)

| Clé `content` | Nb | `level` par défaut | Description |
|---|---|---|---|
| `title` | 1 | 2 | Titre de la section |
| `chart` | 1 | — | Graphique matplotlib |

#### `flowchart_section` — palette bleue

Mise en page : **Titre** / **Logigramme** (tout pleine largeur)

| Clé `content` | Nb | `level` par défaut | Description |
|---|---|---|---|
| `title` | 1 | 2 | Titre de la section |
| `flowchart` | 1 | — | Logigramme simple ou complexe |

> **Recommandation :** toujours utiliser `"flowchart_type": "simple"`. Le mode simple produit des formes Word natives et éditables, sans dépendance supplémentaire. Le mode `"complex"` est réservé aux graphes avec branches multiples et n'est rendu qu'en SVG (HTML) + PNG (Word).

---

### Format des objets de contenu par type singulier

#### `title`

```json
{
  "text": "string",
  "level": 1
}
```

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `text` | string | **oui** | Texte brut affiché |
| `level` | integer | non | `1`, `2` ou `3` — écrase le `level` par défaut du template |

#### `paragraph`

```json
{
  "text": "string"
}
```

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `text` | string | **oui** | Texte brut, sans markdown ni HTML |

#### `bullet_list`

```json
{
  "items": ["string", "string", "string"]
}
```

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `items` | tableau de strings | **oui** | Au moins 1 élément |

#### `image`

```json
{
  "caption": "string",
  "path": null
}
```

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `path` | string ou `null` | **oui** | Chemin vers le fichier image, ou `null` pour afficher un placeholder `[ IMAGE ]` |
| `caption` | string | non | Légende affichée sous l'image |

#### `table`

```json
{
  "headers": ["Colonne 1", "Colonne 2", "Colonne 3"],
  "rows": [
    ["Valeur 1.1", "Valeur 1.2", "Valeur 1.3"],
    ["Valeur 2.1", "Valeur 2.2", "Valeur 2.3"]
  ]
}
```

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `rows` | tableau de tableaux de strings | **oui** | Au moins 1 ligne |
| `headers` | tableau de strings | non | Si absent, pas de ligne d'en-tête |

Le nombre de colonnes est le maximum de `len(headers)` et de `max(len(row) for row in rows)`. Toutes les valeurs sont converties en chaîne.

#### `callout`

```json
{
  "label": "NOTE",
  "text": "string"
}
```

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `text` | string | **oui** | Corps de l'encadré |
| `label` | string | non | Étiquette en gras au-dessus (ex. `"NOTE"`, `"ATTENTION"`, `"ASTUCE"`) — omis = pas d'étiquette |

#### `separator`

```json
{}
```

Aucun champ requis. La clé peut être omise ou avoir pour valeur un objet vide.

#### `chart`

```json
{
  "chart_type": "bar",
  "labels": ["Jan", "Fév", "Mar", "Avr"],
  "datasets": [
    { "label": "Série A", "values": [12, 18, 15, 22] },
    { "label": "Série B", "values": [8, 14, 11, 19] }
  ]
}
```

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `chart_type` | string | **oui** | `"bar"` (barres verticales), `"barh"` (barres horizontales), `"line"` (courbes), `"pie"` (camembert) |
| `labels` | tableau de strings | **oui** | Étiquettes des catégories ou de l'axe X |
| `datasets` | tableau d'objets | **oui** | Au moins 1 dataset — chaque objet a `"label"` (string) et `"values"` (tableau de nombres) |

> Pour `"pie"`, seul le premier dataset est utilisé.

#### `flowchart`

> **Utiliser systématiquement `"flowchart_type": "simple"`** — les formes sont des objets Word natifs (éditables) et la compatibilité est totale. Réserver `"complex"` uniquement quand le graphe contient des branches multiples indispensables.

**Mode simple (recommandé) — séquence linéaire :**

```json
{
  "flowchart_type": "simple",
  "nodes": [
    { "id": "1", "type": "start",    "text": "Début" },
    { "id": "2", "type": "process",  "text": "Traiter la demande" },
    { "id": "3", "type": "decision", "text": "Valide ?" },
    { "id": "4", "type": "io",       "text": "Enregistrer" },
    { "id": "5", "type": "end",      "text": "Fin" }
  ]
}
```

Les nœuds sont reliés dans l'ordre du tableau (pas d'`edges` nécessaire en mode simple).

**Mode complexe — graphe avec branches :**

```json
{
  "flowchart_type": "complex",
  "nodes": [ ... ],
  "edges": [
    { "from": "3", "to": "4", "label": "Oui" },
    { "from": "3", "to": "2", "label": "Non" }
  ]
}
```

**Types de nœuds disponibles :**

| `type` | Forme | Couleur | Usage |
|--------|-------|---------|-------|
| `"start"` | Rectangle arrondi | Vert | Point d'entrée |
| `"end"` | Rectangle arrondi | Rouge | Point de sortie |
| `"process"` | Rectangle | Bleu | Étape de traitement |
| `"decision"` | Losange | Orange | Condition / branchement |
| `"io"` | Parallélogramme | Violet | Entrée / sortie de données |

| Champ | Type | Requis | Contrainte |
|-------|------|--------|------------|
| `flowchart_type` | string | **oui** | `"simple"` (recommandé) ou `"complex"` |
| `nodes` | tableau d'objets | **oui** | Chaque nœud a `"id"` (unique), `"type"` et `"text"` |
| `edges` | tableau d'objets | mode complex | Chaque edge a `"from"`, `"to"` et `"label"` optionnel |

---

### Exemple complet

```json
{
  "title": "Rapport Annuel 2024",
  "blocks": [
    {
      "type": "page",
      "template": "cover_page",
      "sections": [
        {
          "image":     { "caption": "Logo entreprise", "path": null },
          "title":     { "text": "Rapport Annuel 2024", "level": 1 },
          "paragraph": { "text": "Document confidentiel — exercice 2024" }
        }
      ]
    },
    {
      "type": "combined",
      "template": "hero_section",
      "content": {
        "title":     { "text": "Vue d'ensemble", "level": 1 },
        "paragraph": { "text": "Ce rapport présente les résultats consolidés de l'exercice 2024." },
        "image":     { "caption": "Figure 1 : Dashboard exécutif", "path": null }
      }
    },
    {
      "type": "combined",
      "template": "content_with_list",
      "content": {
        "title":       { "text": "Points clés", "level": 2 },
        "paragraph":   { "text": "Les faits marquants de l'exercice :" },
        "bullet_list": { "items": ["Croissance de 15 %", "3 nouveaux marchés", "ISO 27001 obtenu"] }
      }
    },
    {
      "type": "combined",
      "template": "two_column_text",
      "content": {
        "title": [
          { "text": "Avantages", "level": 3 },
          { "text": "Axes d'amélioration", "level": 3 }
        ],
        "paragraph": [
          { "text": "Forte croissance, équipe soudée, produit différenciant." },
          { "text": "Délais de livraison à optimiser, couverture géographique à étendre." }
        ]
      }
    },
    {
      "type": "combined",
      "template": "data_table_section",
      "content": {
        "title": { "text": "Indicateurs financiers", "level": 2 },
        "table": {
          "headers": ["Indicateur", "2023", "2024", "Variation"],
          "rows": [
            ["Chiffre d'affaires", "2,4 M€", "2,76 M€", "+15 %"],
            ["Résultat net",       "380 K€", "456 K€",  "+20 %"]
          ]
        }
      }
    }
  ]
}
```

---

### Guide de génération pour une IA

Pour produire un JSON valide à partir d'une description textuelle, suivre ces règles dans l'ordre :

1. **Choisir le bon template pour chaque section** :

   | Contenu de la section | Template recommandé |
   |---|---|
   | Page de couverture (logo + titre) | bloc `"page"` avec `"cover_page"` |
   | Texte principal + image côte à côte | `"hero_section"` |
   | Texte + liste à puces | `"content_with_list"` |
   | Tableau de données | `"data_table_section"` |
   | Comparaison deux colonnes | `"two_column_text"` |
   | Graphique (courbes, barres, camembert) | `"chart_section"` |
   | Logigramme / processus | `"flowchart_section"` avec `"flowchart_type": "simple"` |

2. **Respecter exactement les clés** de la table de référence du template choisi — aucune clé supplémentaire, aucune clé manquante parmi les requis.

3. **Appliquer la règle des multiples** : si la table indique `Nb = 2`, la valeur est un tableau de 2 objets dans l'ordre gauche→droite.

4. **`level` dans `title`** : ne l'inclure que pour surcharger le niveau par défaut indiqué dans la table (sinon il est injecté automatiquement).

5. **`path: null`** pour toute image dont le fichier n'est pas disponible.

6. **Texte brut uniquement** : pas de markdown (`**gras**`, `# titre`), pas de HTML dans les valeurs de type string.

7. **Toutes les valeurs de cellules de tableau** doivent être des strings (pas des nombres : `"15"` et non `15`).

8. Un document peut mélanger librement des blocs `"page"` et des blocs `"combined"` dans `blocks`.

---

## Créer ses propres templates

### Nouveau template singulier

Créer `mes_templates/singular/mon_bloc.json` :
```json
{
  "id": "mon_bloc",
  "type": "paragraph",
  "font": { "name": "Georgia", "size": 12, "bold": false, "italic": false, "color": "{{text_color}}" },
  "alignment": "justify",
  "spacing": { "before_pt": 8, "after_pt": 8, "line_spacing": 1.5 }
}
```

### Nouveau template combiné

Créer `mes_templates/combined/ma_section.json` avec la structure `color_palette` + `layout.rows[].columns[].components`, en référençant les templates singuliers par leur `id`.

### Nouveau template de page

Créer `mes_templates/pages/ma_page.json` en listant les IDs de templates combinés dans `sections[].combined_template`.

```python
writer = WordDocumentWriter(custom_templates_dir="mes_templates/")
```

---

## Prompt Claude Code — générer la documentation utilisateur d'un projet

Coller ce prompt dans le chat Claude Code d'un autre projet pour qu'il installe `doc_writer` et génère automatiquement une documentation **orientée utilisateur final** (pas développeur) en Word et HTML.

```
Utilise le module `doc_writer` pour générer la documentation utilisateur de ce projet en Word (.docx) et HTML (.html).

**Étapes à faire dans l'ordre :**

1. Crée un environnement virtuel `.venv` s'il n'existe pas, et installe `doc_writer` dedans :
   pip install git+https://github.com/sacha37420/doc_writter.git

2. Lis le README de `doc_writer` pour comprendre le format JSON attendu — en particulier
   la section "Format JSON de document" et le "Guide de génération pour une IA".

3. Explore ce projet (code source, structure, fichiers existants) pour en comprendre
   le fonctionnement et l'usage.

4. Crée le dossier `.doc/` à la racine du projet s'il n'existe pas.

5. Génère un fichier `.doc/documentation.json` qui décrit ce projet selon le format
   `doc_writer`. La documentation doit être entièrement orientée utilisateur final :
   - Ce que fait le projet et à quoi il sert concrètement
   - Comment l'installer et le lancer
   - Comment l'utiliser au quotidien (cas d'usage, exemples concrets)
   - Ce que l'utilisateur peut faire / obtenir
   Ne pas inclure de détails d'implémentation, d'architecture interne ou de concepts
   réservés aux développeurs.

6. Exécute ce script pour produire les deux sorties :

   from doc_writer import WordDocumentWriter, HtmlDocumentWriter
   word = WordDocumentWriter()
   word.write_from_json(".doc/documentation.json")
   word.save(".doc/documentation.docx")
   html = HtmlDocumentWriter()
   html.write_from_json(".doc/documentation.json")
   html.save(".doc/documentation.html")

Livrable : `.doc/documentation.json`, `.doc/documentation.docx`, `.doc/documentation.html`.
```

---

## Structure du projet

```
doc_writer/
├── __init__.py
├── base_writer.py          # Classe abstraite + write_from_json
├── template_loader.py      # Chargement et résolution des JSON
├── word_writer.py          # Générateur Word (.docx)
├── html_writer.py          # Générateur HTML
├── chart_utils.py          # Rendu matplotlib (bar, barh, line, pie)
├── flowchart_utils.py      # Formes Word natives + SVG/PNG logigrammes
└── templates/
    ├── singular/
    │   ├── paragraph.json
    │   ├── bullet_list.json
    │   ├── title.json
    │   ├── image.json
    │   ├── table.json
    │   ├── callout.json
    │   ├── separator.json
    │   ├── chart.json
    │   └── flowchart.json
    ├── combined/
    │   ├── hero_section.json
    │   ├── content_with_list.json
    │   ├── data_table_section.json
    │   ├── cover_section.json
    │   ├── two_column_text.json
    │   ├── chart_section.json
    │   └── flowchart_section.json
    └── pages/
        ├── standard_report_page.json
        └── cover_page.json
examples/
└── example_usage.py
pyproject.toml
README.md
```
