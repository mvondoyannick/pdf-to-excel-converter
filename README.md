# Convertisseur PDF → XLS (web)

Petit utilitaire Flask pour uploader un PDF, le convertir en fichier Excel (.xls) et télécharger le résultat.

Prérequis
- Java installé (tabula-py dépend de Tabula qui requiert Java)
- Python 3.8+

Installation

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

Lancement

```bash
python app.py
# Puis ouvrir http://127.0.0.1:5000
```

Remarques
- Les fichiers uploadés sont enregistrés dans le dossier `uploads` et les sorties dans `converted`.
- Le script `convert_to_xls.py` est utilisé pour la conversion; assurez-vous que `tabula` fonctionne sur votre machine.
