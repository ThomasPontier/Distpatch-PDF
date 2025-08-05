# Dispatch PDF <img width="45" height="45" alt="app1" src="https://github.com/user-attachments/assets/294cc32a-2820-4468-bd5a-a3afb4fdc5c7" />

Application Windows pour analyser des PDF bagages, détecter des pages « escale », prévisualiser, mapper les destinataires et préparer l’envoi d’emails.

Dernière mise à jour : 2025-08-04

## Sommaire
- [Aperçu](#aperçu)
- [Fonctionnalités](#fonctionnalités)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Configuration](#configuration)
- [Utilisation](#utilisation)
- [Architecture technique](#architecture-technique)
- [Build et distribution](#build-et-distribution)
- [Dépannage](#dépannage)
- [FAQ](#faq)
- [Licence](#licence)

## Aperçu
Distpatch PDF facilite le contrôle et la diffusion d’extraits PDF par escale. L’application détecte automatiquement les pages pertinentes, offre un aperçu visuel, applique un mapping escale → destinataires et prépare l’envoi d’emails.

## Fonctionnalités
- Détection automatique des pages « escale » dans un PDF.
- Règles de détection par motif texte (ex. « AAA-Bilan » et présence de mots-clés).
- Liste des escales détectées et navigation par item.
- Prévisualisation PDF intégrée (zoom, pagination).
- Mapping escale → destinataires configurable.
- Préparation d’email(s) par escale avec objet et corps paramétrables (placeholders).
- Gestion de comptes/paramètres via fichiers de configuration.
- Interface utilisateur PySide6.
- Build exécutable Windows via PyInstaller.
- Fichiers de configuration stockés en développement et en mode packagé.

Limitations
- Windows uniquement (intégration et packaging ciblés Windows).

## Prérequis
- Windows 10/11 (x64)
- Python 3.10+ recommandé (voir `pyproject.toml` / `requirements.txt`)
- Dépendances : voir [`requirements.txt`](requirements.txt)

## Installation
Créer l’environnement et installer
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Lancer depuis les sources
```bash
python main.py
```

Créer un exécutable (Windows)
```bash
python build.py
# ou
pyinstaller distpatch_pdf.spec
```

## Configuration
Emplacements par défaut
- Développement : `config/app_config.json`, `config/accounts.json`
- Exécutable : dossier d’application (géré par l’outil de build) ou répertoire dédié utilisateur (selon implémentation)

Fichiers
- `config/app_config.json` : paramètres applicatifs (stopovers, templates, mappings…)
- `config/accounts.json` : comptes/envois (si utilisés par l’application)

Exemple minimal (`app_config.json`)
```json
{
  "version": 1,
  "stopovers": ["CDG", "ORY"],
  "mappings": { "CDG": ["ops.cdg@example.com"] },
  "templates": {
    "subject": "Bilan {STOPOVER}",
    "body": "Bonjour,\nVeuillez trouver le bilan {STOPOVER}.\nCordialement."
  },
  "last_sent": {}
}
```

Placeholders supportés (exemples)
- `{STOPOVER}` : code escale détecté (ex. CDG)
- D’autres variables peuvent être ajoutées selon les besoins du projet.

## Utilisation
Flux rapide
1) Ouvrir l’application (ou `python main.py`)
2) Sélectionner un PDF
3) Cliquer « Analyser »
4) Parcourir la liste des escales détectées
5) Prévisualiser la page correspondante
6) Préparer l’envoi d’emails (à partir du mapping et des templates)

Règles de détection (par défaut)
- Page contenant un motif de type « AAA-Bilan » (AAA = 3 lettres majuscules)
- Présence de mots-clés associés sur la même page (ex. « objectifs »), recherche insensible à la casse
- La logique est centralisée dans `core/detection_engine.py` et peut être adaptée.

Conseils
- Vérifier que les mappings et templates sont cohérents avec les codes escales attendus.
- Valider l’aperçu PDF avant préparation d’email.

## Architecture technique
Entrée principale
- `main.py`

Interface utilisateur (PySide6)
- Fenêtre principale : `ui/pyside_main_window.py`
- Aperçu PDF : `ui/pdf_preview.py`
- Onglets & dialogues : `ui/pyside_*` (mapping, email preview, gestion comptes)
- Style : `ui/style_pyside.qss`
- Composants : `ui/components/*`

Cœur métier
- Détection : `core/detection_engine.py`
- Traitement PDF : `core/pdf_processor.py`
- Rendu PDF : `core/pdf_renderer.py`

Modèles
- Escales : `models/stopover.py`
- Templates : `models/template.py`

Services
- Configuration : `services/config_manager.py`, `services/config_service.py`
- Mapping : `services/mapping_service.py`
- Email escale : `services/stopover_email_service.py`
- Email générique : `services/email_service.py`

Contrôleur
- `controllers/app_controller.py`

Utilitaires
- Fichiers : `utils/file_utils.py`

Documents de cadrage
- `steering/` (product, structure, tech)

## Build et distribution
- Script de build : `build.py` (orchestration PyInstaller)
- Spécification PyInstaller : `distpatch_pdf.spec`
- Assets intégrés : `assets/app.ico`, `assets/app.png`

Commandes
```bash
# développement
python main.py

# build
python build.py
pyinstaller distpatch_pdf.spec
```

Après build
- Un répertoire `dist/` est généré avec l’exécutable et les ressources.
- Vérifier que les fichiers de configuration nécessaires sont présents/initialisés au premier lancement.

## Dépannage
Problèmes courants
- « No module named 'fitz' » : installer PyMuPDF
```bash
pip install PyMuPDF
```
- « Error opening PDF » : vérifier l’intégrité/droits du fichier
- « No stopover pages found » : contrôler les motifs (« AAA-Bilan ») et mots-clés attendus
- Chemins de configuration : s’assurer que `config/*.json` est lisible/valide (JSON)

Logs
- L’application peut écrire des logs (selon configuration). Vérifier la console et/ou les fichiers logs si disponibles.

## FAQ
macOS/Linux sont-ils supportés ?
- L’application est développée et packagée pour Windows.

Où se trouvent les fichiers de configuration ?
- En développement : `config/…`
- En exécutable : selon la stratégie du build, au sein du dossier d’installation/lancement ou d’un répertoire utilisateur.

Peut-on personnaliser l’objet et le corps d’email ?
- Oui via `templates.subject` et `templates.body` (`app_config.json`). Les placeholders comme `{STOPOVER}` sont remplacés au runtime.

Comment étendre les règles de détection ?
- Adapter `core/detection_engine.py` pour ajouter/supprimer des motifs et mots-clés.

## Licence
Usage pratique interne/éducatif, fourni sans garantie.
