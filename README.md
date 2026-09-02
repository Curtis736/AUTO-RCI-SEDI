# AUTO-RCI-SEDI

Application Windows (Python / Tkinter) pour la **génération automatique de fiches Word** (RCI, RE, PV) à partir de mesures Excel, avec insertion d'images, courbes Excel, graphiques dichroïques et remplissage automatique de la **mesure de fin CIT** depuis les données de cyclage thermique.

> Documentation complémentaire : [MODE_EMPLOI.html](MODE_EMPLOI.html) (utilisateur) · [DOCUMENTATION_TECHNIQUE.html](DOCUMENTATION_TECHNIQUE.html) (développeur)

---

## Table des matières

- [Présentation](#présentation)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Démarrage rapide](#démarrage-rapide)
- [Interface — les 4 onglets](#interface--les-4-onglets)
- [Procédure type](#procédure-type)
- [Structure du dossier LT](#structure-du-dossier-lt)
- [Tags du modèle Word](#tags-du-modèle-word)
- [Remplissage automatique CIT](#remplissage-automatique-cit)
- [Configuration](#configuration)
- [Architecture](#architecture)
- [Tests](#tests)
- [Dépannage](#dépannage)
- [Structure du dépôt](#structure-du-dépôt)

---

## Présentation

AUTO_RCI lit un **fichier Excel de mesures** (un onglet par numéro de série), récupère les valeurs associées à chaque SN du lot (LT), ouvre un **modèle Word** (.docx) et produit une fiche par SN dans le dossier de sortie configuré (`RCI/`, `RE/`, `PV/`, etc.).

L'application peut également :

- insérer des **images** depuis les sous-dossiers du LT (JPG, PNG, PDF) ;
- exporter des **courbes Excel** embarquées dans les classeurs du LT ;
- récupérer des **graphiques dichroïques** depuis le classeur « Suivi traitement » ;
- remplir la **fin de mesure CIT** depuis `THERMAL_CYCLING/` si la colonne Excel est vide ;
- convertir en lot des fichiers Word/Excel en **PDF**.

Les fichiers incomplets sont marqués `[INCOMPLET]` dans le nom. Les détails sont dans `Logs/`.

---

## Prérequis

| Élément | Détail |
|---------|--------|
| **OS** | Windows 10/11 |
| **Python** | 3.9+ (3.11+ recommandé) |
| **Microsoft Excel** | Lecture mesures, export courbes (pilotage COM) |
| **Microsoft Word** | Modèles .docx, conversion PDF |
| **Accès réseau** | Dossiers HOI/LT, fichier mesures, modèles formulaires |

> Fermez Excel et Word ouverts manuellement **avant** une génération — le pilotage automatique peut échouer sinon.

---

## Installation

### 1. Cloner ou copier le projet

```bat
git clone https://github.com/Curtis736/AUTO-RCI-SEDI.git
cd AUTO-RCI-SEDI
```

Sur le réseau SEDI, le dossier peut aussi être utilisé directement depuis le partage (ex. `X:\Production\4_Public\DEV (ne pas toucher)\AUTO_RCI_COMPLET`).

### 2. Installer les dépendances Python

Double-cliquer sur **`install_modules.bat`** ou :

```bat
python -m pip install -r requirements.txt
```

Ce script installe notamment : `pywin32`, `python-docx`, `openpyxl`, `Pillow`, `pdf2image`, etc.

### 3. Poppler (PDF → image)

Installé automatiquement dans `tools/poppler/` par `install_modules.bat`. Aucune configuration manuelle par poste.

### 4. Configurer les chemins

Copier le modèle de configuration :

```bat
copy config\paths.json.example config\paths.json
```

Puis adapter `config/paths.json` et `config/fields.json` (voir [Configuration](#configuration)).

### 5. Raccourci (optionnel)

Créer un raccourci bureau vers **`lancer_RCI.bat`**.

---

## Démarrage rapide

1. `install_modules.bat` (première fois)
2. Renseigner LT + formulaire dans l'onglet **Find Data** → **fill**
3. Vérifier les 3 chemins dans le panneau de droite
4. Onglet **Générateur fiches RCI/RE** → **générer document**
5. Récupérer les `.docx` dans `{dossier LT}/{RCI|RE|PV}/`

---

## Interface — les 4 onglets

### 1 — Find Data (remplissage automatique)

| Champ | Action |
|-------|--------|
| **LT** | Numéro de lancement, ex. `LT2600359` |
| **Formulaire** | Ex. `F469 E` (numéro + indice) |
| **Catégorie** | `RE`, `PV` ou `RCI` — définit le dossier de sortie |
| **fill** | Recherche auto du dossier LT, fichier mesures et modèle Word |

### 2 — Paths (chemins manuels)

Si le remplissage auto échoue :

| Bouton | Contenu |
|--------|---------|
| choix chemin mesures | Fichier Excel des mesures |
| choix chemin formulaire | Modèle Word (.docx) |
| choix chemin lancement | Dossier du LT |
| choix chemin Poppler | Laisser vide (auto `tools/poppler/`) |

La config est mémorisée dans `config/paths.json` et `config/fields.json`.

### 3 — Générateur fiches RCI/RE

| Élément | Rôle |
|---------|------|
| format du nom de fichier | Ex. `**REF_SEDI**_SN**SN**_RE` |
| chemin de sortie | Sous-dossier : `RCI`, `PV`, `RE`… |
| réaction aux fichiers préexistants | conserver / dupliquer / remplacer / demander |
| générer document | Lance la génération (tous les SN ou sélection) |

### 4 — Conversion PDF

Conversion par lot des Word/Excel d'un dossier en PDF (filtres par nom de fichier). Indépendant de la génération des fiches.

---

## Procédure type

1. **Find Data** : saisir LT, formulaire, catégorie → **fill** → vérifier les 3 chemins (panneau droit).
2. **Paths** : corriger manuellement si un chemin manque.
3. **Générateur** : vérifier dossier de sortie et format du nom.
4. **générer document** :
   - option 1 : tous les SN du LT ;
   - option 2 : sélection manuelle des numéros de série.
5. Attendre la fin (suivre la zone de log).
6. Récupérer les fichiers dans `{root_work_dir}/{generator_out_path}/`.

### Fichiers générés

| Cas | Signification |
|-----|---------------|
| Fichier normal | Tous les tags remplis, images et courbes trouvées |
| `…[INCOMPLET].docx` | Donnée, image ou courbe manquante — voir `Logs/` |
| Fichier dupliqué | Mode « dupliquer » et fichier déjà existant (suffixe numérique) |

---

## Structure du dossier LT

```
LT2600359/                          ← root_work_dir
├── RE/                             ← sortie rapports d'essai
├── RCI/                            ← sortie fiches RCI
├── PV/                             ← sortie procès-verbaux
├── THERMAL_CYCLING/                ← données CIT (.xlsm, acceptance .XLS)
│   └── RETA-AGS24.134B SN25-20-01 …xlsm
├── DICHRO_GRAND_CHAMP/             ← images dichro
│   └── SN25-20-07.JPG
├── test pression/                  ← courbes Excel
│   └── courbe SN25-20-07.xlsx
└── …
```

### Règles de nommage

| Type | Règle |
|------|-------|
| **Images** | Nom contenant `SN{numéro}`, ex. `SN25-20-07.JPG` |
| **Courbes Excel** | Légende du graphique au format `S/N{numéro}` |
| **Mesures Excel** | Le LT saisi doit exister dans le fichier mesures |
| **Modèle Word** | Tags `**NOM**` (texte), `§§DOSSIER§§` (images) |

---

## Tags du modèle Word

| Format | Usage |
|--------|-------|
| `**TAG**` | Valeur texte lue dans l'Excel de mesures |
| `$$TAG$$` | Colonne de tableau (séries de mesures) |
| `§§NOM_DOSSIER§§` | Images de `{root_work_dir}/NOM_DOSSIER/` filtrées par SN |
| `§§NOM_DOSSIER_01§§` | Même principe avec index (`_01`, `_02`…) |
| `§§DICHRO_GRAPH§§` | Graphique dichro exporté depuis « Suivi traitement » |

Les courbes Excel référencées dans un tag `§§…§§` sont exportées via COM Excel. Les PDF passent par Poppler.

---

## Remplissage automatique CIT

Si la colonne **CIT** / **FIN_CIT** / **fin_cit** est **vide** dans l'Excel de mesures, AUTO_RCI tente de la remplir depuis `{root_work_dir}/THERMAL_CYCLING/`.

### Source des données

1. **Prioritaire** : fichiers `.xlsm` / `.xlsx`, feuille `data_traitee`, en-têtes `REF … SN…` (ligne 3), **dernière valeur numérique** de la colonne.
2. **Fallback** : fichiers `*acceptance*.XLS` (OptoTest).

### Valeur produite

`abs(dernière mesure)`, arrondie à **3 décimales** (format point : `0.052`).

### Correspondance SN

Le matching utilise **référence + numéro de série** pour éviter les collisions entre plans (ex. SN144 sur 23.157 vs 23.159). Formats SN supportés : `144`, `25-20-01`, `SN25-20-01`.

### Tag Word à utiliser

```
**FIN_CIT**
```

Alias acceptés dans l'Excel : `CIT`, `FIN_CIT`, `fin_cit`.

### Performance (lots volumineux)

Pour 200+ ou 500+ mesures, l'indexation CIT lit chaque fichier `THERMAL_CYCLING` **une seule fois** puis fait des lookups en mémoire. **`openpyxl` est fortement recommandé** (installé via `requirements.txt`) — sans lui, la lecture repasse par Excel COM (plus lent et fragile).

Logs à surveiller :

```
[CIT] Index construit: N mesure(s) de fin
[CIT] N mesure(s) auto-remplie(s), X SN sans donnée CIT
```

---

## Configuration

### `config/paths.json`

| Clé | Description |
|-----|-------------|
| `measurement_file` | Fichier Excel de mesures |
| `template_file_path` | Modèle Word (.docx) |
| `root_work_dir` | Dossier LT (images, THERMAL_CYCLING, sorties) |
| `generator_out_path` | Sous-dossier de sortie (`RE`, `RCI`, `PV`…) |
| `dichroic_root` | Racine de recherche « Suivi traitement » (optionnel) |
| `dichroic_workbook` | Chemin explicite du classeur dichro (sinon recherche auto) |
| `dichroic_filename_hint` | Mot-clé du nom de fichier, ex. `suivi traitement` |
| `poppler_path` | Laisser `null` (auto) sauf cas particulier |

Modèle vierge : [`config/paths.json.example`](config/paths.json.example).

### `config/fields.json`

| Clé | Description |
|-----|-------------|
| `LT` | Numéro de lancement |
| `form` | Formulaire, ex. `F469 E` |
| `out_file_category` | `RE`, `PV` ou `RCI` |
| `generator_out_filename` | Modèle du nom généré, ex. `**REF_SEDI**_SN**SN**_RE` |
| `file_reaction_mode` | `remplacer`, `conserver`, `dupliquer`, `demer` |
| `num_plan` | Numéro de plan |
| `pdf_input_path` | Dossier source pour conversion PDF par lot |

> Sur un dossier réseau partagé, le dernier utilisateur écrase `config/` — reconfigurer via **Find Data** si besoin.

---

## Architecture

```
Interface.py              ← point d'entrée (lancer_RCI.bat)
├── PathsInterface        ← sélection des chemins
├── FieldFillerInterface  ← auto-remplissage depuis LT
├── DocumentGenerator     ← orchestration génération
│   ├── ValueFetcher      ← lecture Excel → Container par SN
│   │   └── CitMeasurementFetcher  ← indexation CIT / THERMAL_CYCLING
│   ├── Writer            ← remplacement tags Word + images
│   ├── CustomImageTagProcessor  ← tags §§...§§
│   ├── ExcelController   ← COM Excel (courbes, dichro)
│   └── ImageReader       ← cache images, PDF→PNG
├── PdfGenerator          ← conversion PDF par lot
├── poppler_setup         ← installation auto Poppler
└── Settings / Log / FileFinder
```

### Flux de génération

1. Charger la config (`Settings`, `paths.json`, `fields.json`).
2. Lire l'Excel → un **container** par SN (`ValueFetcher.FetchValuesInExcel()`).
3. Auto-remplir CIT si colonnes vides (`CitMeasurementFetcher`).
4. Filtrer les SN (tous ou sélection).
5. Pour chaque SN : ouvrir le modèle Word, remplacer tags, insérer images/courbes, sauvegarder.

---

## Tests

### Tests unitaires

```bat
lancer_tests.bat
```

ou :

```bat
python -m unittest discover -s tests -v
```

Couverture : CIT, FileFinder, Log, WordController, régressions SN/référence.

### Tests terrain (données LT réelles)

Le dossier **`test-unitaire/`** (~230 Mo) n'est **pas versionné** — à placer localement pour les tests terrain.

```bat
tests_terrain\lancer_demo_tests.bat    ← démo : 8 succès + 5 échecs volontaires
tests_terrain\lancer_tests_terrain.bat ← suite complète
```

Profil LT de référence : `tests_terrain/lt_profile.json`.

Résultat attendu de la démo :

```
Étape 1 : 8/8 OK
Étape 2 : 5/5 échoués → OK (démo valide)
```

Documentation détaillée : [tests_terrain/DOCUMENTATION_TESTS.html](tests_terrain/DOCUMENTATION_TESTS.html).

---

## Dépannage

| Problème | Solution |
|----------|----------|
| Python introuvable | Réinstaller Python avec « Add to PATH », relancer `install_modules.bat` |
| Aucun SN pour le LT | Vérifier fichier mesures et présence du LT dans l'Excel |
| Formulaire non trouvé | Vérifier format `Fxxx A` et accès dossier formulaires réseau |
| Images non insérées | Nom `SN…`, extension (.JPG), sous-dossier du tag Word |
| Courbes Excel absentes | Fermer Excel, vérifier légende `S/N…`, relancer |
| CIT non remplie | SN absent de `THERMAL_CYCLING`, ou REF différente ; voir logs `[CIT]` |
| Plantage 200+ mesures | Installer `openpyxl` ; fermer Excel/Word avant le lot |
| Config écrasée | Reconfigurer via **Find Data** sur dossier réseau partagé |
| `[INCOMPLET]` | Consulter `Logs/problems.log` et `Logs/verbose.log` |

---

## Structure du dépôt

```
AUTO-RCI-SEDI/
├── Interface.py              # Application principale
├── DocumentGenerator.py      # Génération des fiches
├── ValueFetcher.py           # Lecture Excel
├── CitMeasurementFetcher.py  # Récupération fin CIT
├── Writer.py                 # Manipulation Word
├── ExcelController.py        # Pilotage Excel COM
├── config/                   # Configuration utilisateur
├── tests/                    # Tests unitaires
├── tests_terrain/            # Tests sur LT réels
├── tools/                    # Poppler (auto-installé)
├── lancer_RCI.bat            # Lancement application
├── install_modules.bat       # Installation dépendances
├── MODE_EMPLOI.html          # Guide utilisateur
└── DOCUMENTATION_TECHNIQUE.html
```

---

## Licence / usage

Outil interne SEDI — génération de documentation de contrôle à partir de données métier (Excel, cyclage thermique, images).
