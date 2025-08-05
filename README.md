# Dispatch PDF <img width="45" height="45" alt="app1" src="https://github.com/user-attachments/assets/294cc32a-2820-4468-bd5a-a3afb4fdc5c7" />

Application Windows pour analyser des PDF bagages, détecter des pages « escale », prévisualiser, mapper les destinataires et préparer l’envoi d’emails.

Dernière mise à jour : 2025-08-04

## Sommaire
- [Aperçu](#aperçu)
- [Fonctionnalités](#fonctionnalités)
- [Installation](#installation)
- [Configuration](#configuration)
- [Utilisation](#utilisation)
- [Licence](#licence)

## Aperçu
Distpatch PDF facilite le contrôle et la diffusion d’extraits PDF par escale. L’application détecte automatiquement les pages pertinentes, offre un aperçu visuel, applique un mapping escale → destinataires et prépare l’envoi d’emails.

## Fonctionnalités
- Détection automatique des pages « escale » dans un PDF.
- Règles de détection par motif texte (ex. « AAA-Bilan » et présence de mots-clés).
- Liste des escales détectées et navigation par item.
- Prévisualisation PDF intégrée (zoom, pagination).
- Mapping escale → destinataires configurable.
- Préparation d’email(s) par escale avec objet et corps paramétrables.
- Interface utilisateur PySide6.
- Build exécutable Windows via PyInstaller.
- Fichiers de configuration stockés en développement et en mode packagé.



## Installation
# Téléchargez et lancez le fichier executable :




# Ou depuis le code source :
Créer l’environnement et installer les dépendances
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Lancer depuis les sources
```bash
python main.py
```

Créer un exécutable 
```bash
python build.py
# ou
pyinstaller distpatch_pdf.spec
```

## Configuration
En mode packagé, le dossier de configuration est stocké dans %APPDATA%/Roaming

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





## Licence
Usage pratique interne/éducatif, fourni sans garantie.
