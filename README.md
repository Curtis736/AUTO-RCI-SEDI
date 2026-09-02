# AUTO-RCI-SEDI

Génération automatique de fiches RE/RCI à partir de mesures Excel, données THERMAL_CYCLING et modèles Word.

## Démarrage rapide

1. Installer les dépendances : `install_modules.bat`
2. Copier `config/paths.json.example` vers `config/paths.json` et adapter les chemins
3. Lancer : `lancer_RCI.bat`

## Documentation

- Mode d'emploi utilisateur : `MODE_EMPLOI.html` (ou `ouvrir_mode_emploi.bat`)
- Documentation technique : `DOCUMENTATION_TECHNIQUE.html`

## Tests

```bat
lancer_tests.bat
```

Les tests unitaires CIT utilisent des fichiers temporaires. Le dossier `test-unitaire/` (données LT réelles) reste **local** et n'est pas versionné.

## Prérequis

- Windows, Python 3.11+
- Microsoft Word et Excel
- `openpyxl` recommandé pour la lecture CIT (installé via `requirements.txt`)
