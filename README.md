# Planning Vacations SEM – DRMCO

## Lancement local

1. Installer Python 3.10 ou plus récent.
2. Ouvrir un terminal dans ce dossier.
3. Installer les dépendances :

```bash
pip install -r requirements.txt
```

4. Lancer l'application :

```bash
streamlit run app.py
```

## Première connexion

- Identifiant : `chef`
- Mot de passe initial : `ChangeMe123!`

**Changez immédiatement ce mot de passe** dans « Mon compte ».

## Comptes à créer

Depuis « Administration », le Chef de service crée :
- le compte du Contrôleur ;
- les cinq comptes des prévisionnistes en rotation.

## Règles intégrées dans ce prototype

- Le Chef de service et le Contrôleur sont affichés comme réguliers du lundi au vendredi.
- Les cinq prévisionnistes peuvent s'inscrire eux-mêmes à une vacation du matin.
- Deux vacations maximum sont acceptées par jour.
- La période affichée couvre cinq semaines à partir du lundi de la semaine courante.

## Important avant publication en ligne

Cette première version utilise SQLite, pratique pour les essais sur un seul ordinateur.
Pour un usage simultané et durable sur Internet, il faut remplacer SQLite par une base PostgreSQL managée (par exemple Supabase ou un service PostgreSQL compatible), ajouter une authentification renforcée et vérifier les règles de sécurité/hébergement de la DGM. Ne publiez pas cette version telle quelle avec des données professionnelles sensibles.
