# Distpatch-PDF

Application Windows pour analyser des PDF, détecter des pages « escale », prévisualiser et préparer l’envoi par email via Outlook.

Dernière mise à jour : 2025-08-03

## Sommaire

- [Introduction](#introduction)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Configuration](#configuration)
- [Utilisation](#utilisation)
- [Développement](#développement)
- [FAQ](#faq)

## Introduction

But
- Détecter automatiquement dans un PDF les pages contenant un code escale « AAA-Bilan » et le mot « objectifs ».
- Lister les escales détectées et prévisualiser la page.
- Préparer l’envoi d’emails via Outlook avec des modèles personnalisables.

Cas d’usage
- Contrôle rapide d’un rapport PDF multi-escales.
- Préparation d’emails par escale selon un mapping escale → destinataires.

Limitations
- Windows uniquement (intégration Outlook via pywin32). macOS/Linux non supportés pour l’envoi d’emails.

## Prérequis

- Windows 10/11 (x64)
- Python 3.8+ (recommandé)
- Outlook configuré (profil actif) pour l’envoi
- Dépendances : voir [`requirements.txt`](requirements.txt)

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
```

## Configuration

Emplacement du fichier
- Développement : `config/app_config.json`
- Exécutable : `%APPDATA%\Distpatch-PDF\config\app_config.json` (créé au besoin)

Exemple minimal
```json
{
  "version": 1,
  "stopovers": ["CDG"],
  "mappings": { "CDG": ["ops.cdg@example.com"] },
  "templates": { "subject": "Bilan {STOPOVER}", "body": "Bonjour,\nBilan {STOPOVER}." },
  "last_sent": {}
}
```


## Utilisation

Flux rapide
1) Ouvrir l’application (ou `python main.py`)  
2) Sélectionner un PDF  
3) Cliquer « Analyser »  
4) Parcourir les escales détectées et prévisualiser  
5) Préparer l’envoi Outlook (Windows)

Règles de détection
- Code « AAA-Bilan » (AAA = 3 lettres majuscules)
- Mot « objectifs » présent sur la même page (casse insensible)

Erreurs fréquentes
- « No module named 'fitz' » : installer PyMuPDF
```bash
pip install PyMuPDF
```
- « Error opening PDF » : vérifier l’intégrité/droits du fichier
- « No stopover pages found » : vérifier le motif exact « AAA-Bilan » et la présence de « objectifs »

## Développement

Points clés
- Entrée : [`main.py`](main.py)
- UI : PySide6 (fenêtre principale : [`ui.pyside_main_window.py`](ui/pyside_main_window.py))
- Détection/rendu PDF : [`core/`](core/)
- Configuration unifiée : [`services.config_manager`](services/config_manager.py)

Commandes
```bash
python main.py
python build.py
pyinstaller distpatch_pdf.spec
```

## FAQ

macOS/Linux sont-ils supportés ?
- Non pour l’envoi d’emails (pywin32/Outlook est Windows-only). L’application est ciblée Windows.

Où se trouve la configuration packagée ?
- `%APPDATA%\Distpatch-PDF\config\app_config.json`.

Peut-on personnaliser l’objet/le corps d’email ?
- Oui via `templates.subject` et `templates.body` dans `app_config.json` (placeholder : `{STOPOVER}`).



Licence : usage éducatif et pratique, sans garantie.
