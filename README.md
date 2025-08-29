# Guide très simple pour utiliser le logiciel AUTO RCI

---

## 1. C'est quoi ce logiciel ?

Ce logiciel sert à créer automatiquement des documents (rapports) à partir de tableaux Excel et de modèles Word. Il est utilisé dans l'industrie pour faire des comptes-rendus de contrôle.

---

## 2. Ce qu'il faut avant de commencer

- **Un ordinateur avec Windows** (c'est le système le plus courant)
- **Microsoft Word** (pour ouvrir les documents Word)
- **Microsoft Excel** (pour ouvrir les tableaux Excel)
- **Python** (un programme spécial, demandez à quelqu'un de l'installer si ce n'est pas déjà fait)

---

## 3. Installer le logiciel

1. **Téléchargez le dossier du logiciel** (demandez à quelqu'un si besoin).
2. **Ouvrez le dossier** où il y a le logiciel.
3. **Double-cliquez** sur le fichier qui s'appelle `lancer_RCI.bat`.
   - Si rien ne se passe, demandez de l'aide à quelqu'un.

---

## 4. Ouvrir le logiciel

- Après avoir double-cliqué, une fenêtre va s'ouvrir avec plusieurs onglets en haut (exemple : "Paths", "Générateur fiches RCI/RE").

---

## 5. Préparer le logiciel

### a) Dire où sont vos fichiers

- **Chemin du dossier de travail** : Cliquez sur l'onglet "Paths" et indiquez où est votre dossier principal (là où sont vos fichiers).
- **Fichier Excel** : Indiquez où est votre tableau Excel (demandez le chemin si vous ne savez pas).
- **Modèle Word** : Indiquez où est votre modèle Word (fichier .docx).

### b) Remplir les informations de base

- **Numéro de lancement (LT)** : C'est un code qui commence par "LT" (exemple : LT2500414).
- **Formulaire** : C'est le nom du modèle Word (exemple : F458 E).
- **Catégorie** : Choisissez le type de rapport (RE, PV, ou RCI).

---

## 6. Créer un document

1. Cliquez sur l'onglet "Générateur fiches RCI/RE".
2. Cliquez sur le bouton "générer document".
3. Le logiciel va créer un ou plusieurs fichiers Word automatiquement, en utilisant vos données Excel et le modèle Word.

---

## 7. Où trouver les documents créés ?

- Le logiciel vous dit dans quel dossier il a mis les nouveaux fichiers.
- Si un message d'erreur apparaît, lisez-le : il peut vous dire si un dossier ou un fichier manque.

---

## 8. Conseils importants

- **Ne changez pas le nom des dossiers ou fichiers sans demander à quelqu'un.**
- **Si le logiciel affiche une erreur, lisez le message et demandez de l'aide si besoin.**
- **Gardez toujours une copie de vos fichiers originaux.**
- **Fermez Word et Excel avant de lancer le logiciel pour éviter les problèmes.**

---

## 9. Problèmes fréquents et solutions

- **Le logiciel ne démarre pas** : Vérifiez que Python, Word et Excel sont installés.
- **Erreur "le dossier n'existe pas"** : Vérifiez que le dossier existe bien et que vous avez le droit d'y accéder.
- **Erreur "tag non remplacé"** : Vérifiez que votre fichier Excel contient bien toutes les colonnes nécessaires (surtout la colonne "SN").

---

## 10. Besoin d'aide ?

- Demandez à votre service informatique ou à un collègue qui connaît bien l'ordinateur.
- Gardez ce guide à côté de vous pour suivre les étapes.

---

**Rappel :**
Il n'y a pas de question bête. Si vous êtes bloqué, demandez de l'aide !
Ce logiciel est là pour vous faire gagner du temps, pas pour vous compliquer la vie.

# AUTO RCI - Générateur Automatique de Documents RCI/RE

## Description

AUTO RCI est un logiciel de génération automatique de documents RCI (Rapport de Contrôle Interne) et RE (Rapport d'Essai) à partir de données Excel et de templates Word. Il permet de créer des documents standardisés en remplaçant automatiquement des tags dans les templates par les valeurs correspondantes extraites des fichiers Excel de mesures.

## Prérequis

### Logiciels requis
- **Python 3.7+** installé sur votre système
- **Microsoft Excel** (pour la lecture des fichiers de mesures)
- **Microsoft Word** (pour la génération des documents)

### Installation des dépendances Python

1. Ouvrez un terminal dans le dossier du projet
2. Exécutez la commande suivante pour installer les dépendances :

```bash
pip install -r requirements.txt
```

Les dépendances principales sont :
- `pywin32` : Interface avec Microsoft Office
- `Pillow` : Traitement d'images
- `numpy` : Calculs numériques
- `python-docx` : Manipulation de documents Word
- `docx2pdf` : Conversion Word vers PDF
- `pdf2image` : Conversion PDF vers images
- `pypdf` : Manipulation de fichiers PDF

## Lancement du logiciel

### Méthode 1 : Double-clic
Double-cliquez sur le fichier `lancer_RCI.bat` pour démarrer l'application.

### Méthode 2 : Ligne de commande
```bash
python Interface.py
```

## Configuration initiale

### 1. Configuration des chemins

Lors du premier lancement, vous devez configurer trois chemins essentiels :

1. **Dossier racine de travail** : Le dossier principal contenant vos données
2. **Fichier Excel des mesures** : Le fichier Excel contenant vos données de mesures
3. **Template Word** : Le fichier modèle Word pour générer les documents

### 2. Configuration des champs

Dans l'onglet "Champs", configurez :
- **LT** : Numéro de lancement (ex: LT2500414)
- **Form** : Numéro de formulaire (ex: F458 E)
- **Format du nom de fichier** : Format pour nommer les documents générés
- **Réaction aux fichiers préexistants** : Que faire si un fichier existe déjà

### 3. Utilisation de "Find Data" (Recherche automatique)

L'onglet "Find Data" permet de remplir automatiquement les chemins et configurations en se basant sur la structure de dossiers :

1. **Saisissez le numéro de lancement** (ex: LT2500414)
2. **Saisissez le formulaire** (ex: F458 E)
3. **Sélectionnez la catégorie** de fichier (RE, PV, ou RCI)
4. **Cliquez sur "fill"** pour lancer la recherche automatique

Le logiciel va alors :
- **Rechercher le dossier LT** dans `X:/Tracabilite` (structure attendue : `X:/Tracabilite/RCT°-AGS23.159B/LT2500414`)
- **Trouver le fichier de mesures** dans le dossier parent (ex: `Mesures 23.159.xlsx`)
- **Localiser le template Word** dans `X:/Qualite/4_Public/A disposition/DOSSIER SMI/Formulaires/B3-PRODUCTION`
- **Configurer automatiquement** tous les chemins et formats de noms

**Structure de dossiers attendue :**

```
X:/Tracabilite/
└── RCT°-AGS23.159B/
    ├── LT2500414/          ← Dossier de travail (recherché par nom exact)
    └── Mesures 23.159.xlsx ← Fichier de mesures (recherché par pattern "mesures")

X:/Qualite/4_Public/A disposition/DOSSIER SMI/Formulaires/B3-PRODUCTION/
└── [Sous-dossiers]/
    └── F458 ind E - PV Cordon OHA.docx ← Template (recherché par "F458" + "ind E")
```

**Exemples de noms acceptés :**
- **Fichiers de mesures** : "Mesures 23.159.xlsx", "mesures_2024.xlsx", "MESURES_AG23.xlsx"
- **Templates** : "F458 ind E - PV.docx", "F458_ind_E_PV.docx", "F458 ind E PV Cordon.docx"

## Structure des fichiers Excel

### Format attendu

Le fichier Excel des mesures doit contenir :

1. **En-têtes avec tags** : La première ligne contient les noms des colonnes avec des tags entre `**`
2. **Données par SN** : Chaque ligne correspond à un numéro de série (SN) différent

### Exemple de structure Excel

| SN | **RL1** | **RL2** | **IL_BEFORE** | **IL_AFTER** | **DATE_TEST** |
|----|---------|---------|---------------|--------------|---------------|
| 001 | -45.2 | -42.1 | -0.8 | -0.9 | 2024-01-15 |
| 002 | -48.3 | -44.7 | -0.7 | -0.8 | 2024-01-16 |

### Tags spéciaux

- `**SN**` : Numéro de série (automatiquement rempli)
- `**DATE**` : Date actuelle
- `**LT**` : Numéro de lancement
- `**NUM_PLAN**` : Numéro de plan
- `**RL1**`, `**RL2**` : Mesures de perte de retour (dB)
- `**IL_BEFORE**`, `**IL_AFTER**` : Mesures de perte d'insertion (dB)

### Spécifications de mesures

Vous pouvez ajouter des spécifications pour valider les mesures :

- Format : `<= valeur` ou `>= valeur`
- Exemple : `<= -40` pour RL (perte de retour ≤ -40 dB)
- Exemple : `>= -3` pour IL (perte d'insertion ≥ -3 dB)

## Structure des templates Word

### Tags de remplacement

Dans votre template Word, utilisez des tags entre `**` pour les remplacements automatiques :

```
Document de test pour le SN**SN**
Date de test : **DATE**
Numéro de lancement : **LT**
Mesure RL1 : **RL1** dB
Mesure IL : **IL_BEFORE** dB
```

### Tags pour les tableaux

Pour remplir des tableaux automatiquement, utilisez des tags avec `$$` :

```
| SN | RL1 | RL2 | IL |
| $$SN$$ | $$RL1$$ | $$RL2$$ | $$IL$$ |
```

### Tags pour les images

Pour insérer des images automatiquement, utilisez des tags avec `§§` :

```
Graphique des mesures :
§§GRAPH_RL§§
```

### Images automatiques

Le logiciel peut automatiquement :
- Chercher des images dans des dossiers spécifiques
- Les associer aux SN correspondants
- Les insérer dans les documents

## Utilisation du logiciel

### 1. Configuration rapide avec "Find Data"

1. Allez dans l'onglet "Find Data"
2. Saisissez le numéro de lancement (ex: LT2500414)
3. Saisissez le formulaire (ex: F458 E)
4. Sélectionnez la catégorie de fichier (RE, PV, ou RCI)
5. Cliquez sur "fill" pour configurer automatiquement tous les chemins
6. Vérifiez que tous les chemins ont été trouvés dans la fenêtre de sortie

### 2. Génération d'un document unique

1. Allez dans l'onglet "Générateur fiches RCI/RE"
2. Cliquez sur "générer document"
3. Sélectionnez le SN pour lequel générer le document
4. Le document sera créé dans le dossier de sortie configuré

### 3. Génération en lot

1. Dans l'onglet "Générateur fiches RCI/RE"
2. Sélectionnez plusieurs SN
3. Le logiciel générera un document pour chaque SN sélectionné

### 4. Gestion des fichiers existants

Le logiciel propose plusieurs options :
- **Conserver l'original** : Ne pas écraser les fichiers existants
- **Dupliquer** : Créer une copie avec un suffixe
- **Remplacer** : Écraser le fichier existant
- **Demander** : Demander confirmation pour chaque fichier

## Fonctionnalités avancées

### Recherche automatique de données

Le logiciel peut automatiquement :
- **Rechercher les dossiers LT** dans la structure de dossiers standardisée
- **Localiser les fichiers de mesures** basés sur des patterns de noms
- **Trouver les templates Word** selon le numéro de formulaire et l'indice
- **Configurer automatiquement** tous les chemins et paramètres

#### Algorithme de recherche détaillé

**1. Recherche du dossier LT :**
- **Dossier racine** : `X:/Tracabilite`
- **Profondeur maximale** : 3 niveaux de sous-dossiers
- **Pattern de recherche** : Le numéro de lancement exact (ex: "LT2500414")
- **Résultat** : Chemin complet vers le dossier LT

**2. Recherche du fichier de mesures :**
- **Emplacement** : Dossier parent du dossier LT
- **Pattern de recherche** : Nom contenant "mesures" (insensible à la casse)
- **Exemples acceptés** : "Mesures 23.159.xlsx", "mesures_2024.xlsx"
- **Résultat** : Chemin complet vers le fichier Excel

**3. Recherche du template Word :**
- **Dossier racine** : `X:/Qualite/4_Public/A disposition/DOSSIER SMI/Formulaires/B3-PRODUCTION`
- **Pattern numéro** : "F" + numéro (ex: "F458" pour formulaire "F458 E")
- **Pattern indice** : "ind" + lettre (ex: "ind E" pour formulaire "F458 E")
- **Recherche** : Récursive dans tous les sous-dossiers
- **Résultat** : Chemin complet vers le fichier .docx

**4. Configuration automatique :**
- **Dossier de sortie** : Nom de la catégorie (RE, PV, ou RCI)
- **Format de nom** : `**REF_SEDI**_SN**SN**_[CATEGORIE]`
- **Mise à jour** : Tous les champs de configuration

### Validation automatique des mesures

Le logiciel vérifie automatiquement :
- Les plages de valeurs typiques pour les mesures en dB
- Les spécifications définies dans le fichier Excel
- La cohérence des données

### Gestion des erreurs

- **Logs détaillés** : Tous les événements sont enregistrés dans les fichiers de log
- **Gestion des timeouts** : Gestion robuste des connexions Excel/Word
- **Récupération d'erreurs** : Tentatives automatiques en cas d'échec

### Personnalisation

- **Thèmes** : Personnalisation des couleurs de l'interface
- **Chemins relatifs** : Support des chemins réseau et relatifs
- **Formats de noms** : Personnalisation des noms de fichiers générés

## Fichiers de configuration

### config/fields.json
```json
{
    "LT": "LT2500414",
    "form": "F458 E",
    "out_file_category": "PV",
    "generator_out_filename": "**REF_SEDI**_SN**SN**_PV",
    "file_reaction_mode": "remplacer"
}
```

### config/paths.json
```json
{
    "generator_out_path": "PV",
    "measurement_file": "X:\\Tracabilite\\RCT°-AGS23.159B\\Mesures 23.159.xlsx",
    "template_file_path": "X:/Production/2_Workgroup/Partage Méthode/Modes Opératoires/AGS/A relire/F458 - ind E - PV Cordon OHA.docx",
    "root_work_dir": "X:\\Tracabilite\\RCT°-AGS23.159B\\LT2500414"
}
```

## Points d'attention importants

### ⚠️ Limitations et contraintes

**1. Structure de dossiers obligatoire :**
- Le dossier LT doit être dans `X:/Tracabilite` à maximum 3 niveaux de profondeur
- Le fichier de mesures doit contenir le mot "mesures" dans son nom
- Les templates Word doivent suivre le format "F[numero] ind [lettre]"

**2. Données Excel :**
- **Colonne SN obligatoire** : Le fichier Excel doit contenir une colonne avec le tag `**SN**`
- **Valeurs manquantes** : Le logiciel gère les valeurs vides mais affiche des avertissements
- **Format des mesures** : Les valeurs en dB doivent être numériques (pas de texte)

**3. Mesures en dB :**
- **Return Loss (RL)** : Généralement négatives, plage typique -70 dB à -10 dB
- **Insertion Loss (IL)** : Généralement négatives, plage typique -3 dB à 0 dB
- **Valeurs positives** : Affichent un avertissement mais ne bloquent pas la génération

**4. Spécifications :**
- **Format supporté** : `<= valeur` ou `>= valeur` uniquement
- **Format non supporté** : `=`, `<`, `>`, `!=`
- **Valeurs décimales** : Utilisez des points (pas des virgules)

**5. Templates Word :**
- **Format requis** : Fichiers .docx uniquement
- **Tags sensibles** : Respectez la casse (`**SN**` ≠ `**sn**`)
- **Images** : Formats supportés : PNG, JPG, JPEG, BMP, TIFF, PDF, XLSX

**6. Gestion des erreurs :**
- **Fichiers existants** : Le logiciel peut écraser les fichiers existants
- **Chemins réseau** : Supportés mais peuvent être plus lents
- **Permissions** : Vérifiez les droits d'écriture dans les dossiers de sortie

### 🔍 Validation automatique

Le logiciel vérifie automatiquement :
- **Plages de valeurs** pour les mesures en dB
- **Format des numéros de lancement** (doit commencer par "LT")
- **Cohérence des données** entre le fichier Excel et les paramètres
- **Existence des fichiers** et dossiers requis

### 📊 Gestion des données manquantes

**Valeurs par défaut automatiques :**
- **Date** : Date actuelle
- **Numéro de plan** : "00.000"
- **LT** : Extrait du chemin du dossier de travail

**Tags critiques manquants :**
- **SN** : Bloque la génération
- **Autres tags** : Remplacés par des chaînes vides

## Dépannage

### Problèmes courants

1. **Excel ne s'ouvre pas**
   - Vérifiez que Microsoft Excel est installé
   - Redémarrez l'application

2. **Word ne s'ouvre pas**
   - Vérifiez que Microsoft Word est installé
   - Vérifiez que le template n'est pas corrompu

3. **Erreurs de chemins**
   - Vérifiez que tous les chemins sont accessibles
   - Utilisez des chemins absolus si nécessaire

4. **Tags non remplacés**
   - Vérifiez la syntaxe des tags (`**TAG**`)
   - Vérifiez que les données existent dans le fichier Excel

5. **"Find Data" ne trouve pas les fichiers**
   - Vérifiez que la structure de dossiers respecte le format attendu
   - Vérifiez que le numéro de lancement est correct
   - Vérifiez que le formulaire existe dans le dossier des formulaires
   - Vérifiez les permissions d'accès aux dossiers
   - **Pour le dossier LT** : Vérifiez qu'il est dans `X:/Tracabilite` et à maximum 3 niveaux de profondeur
   - **Pour le fichier de mesures** : Vérifiez qu'il contient le mot "mesures" dans son nom
   - **Pour le template** : Vérifiez qu'il contient "F[numero]" et "ind [lettre]" dans son nom

### Logs

Les logs sont disponibles dans le dossier `Logs/` :
- `normal.log` : Messages généraux
- `problems.log` : Erreurs et avertissements
- `verbose.log` : Messages détaillés

## Support

Pour toute question ou signalement de bug, veuillez contacter le support technique.

---

**Version** : 26/03/2024  
**Développé pour** : Génération automatique de documents RCI/RE 