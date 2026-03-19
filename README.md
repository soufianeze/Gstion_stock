# CVE App — Gestion de Stock

Application web de gestion de stock, déployée sur Railway.

## Lancer en local

```bash
pip install flask openpyxl
python serveur.py
```

Ouvrir : http://localhost:5000

## Déploiement Railway

- `requirements.txt` : dépendances Python
- `Procfile` : commande de démarrage via gunicorn
