# Avibase Morocco Exact Excel Scraper

Scraper Python basé sur `requests` et `BeautifulSoup` pour collecter la checklist Avibase du Maroc, enrichir les espèces via Oiseaux.net / the-birds.net, puis exporter un fichier Excel multi-feuilles sans navigateur graphique.

Le projet conserve la logique métier existante autour de :
- la checklist Maroc Avibase
- le matching exact par nom scientifique
- les synonymes
- les noms vernaculaires
- les distributions et régions du Maroc
- l'enrichissement Oiseaux.net / the-birds.net
- l'export Excel multi-feuilles

## Fonctionnalités

- Scraping de la checklist Avibase Maroc
- Extraction des fiches espèces
- Extraction des synonymes
- Extraction des vernacular names
- Extraction des distributions et du type de présence
- Extraction des régions du Maroc
- Enrichissement Oiseaux.net / the-birds.net par nom scientifique exact
- Export Excel avec feuilles structurées
- Reprise automatique via `checkpoints/progress.json`
- Sauvegardes partielles régulières dans `output/`
- Logs détaillés dans `logs/scraper.log`
- Notification email de fin ou d'échec critique
- Compatible Railway

## Structure

```text
.
├── avibase_exact_excel_scraper_FR_OISEAUX.py
├── requirements.txt
├── README.md
├── railway.json
├── .gitignore
├── .env.example
├── logs/
├── checkpoints/
├── output/
└── utils/
    ├── __init__.py
    ├── mailer.py
    ├── checkpoint.py
    ├── logger_config.py
    └── retry_utils.py
```

Les dossiers `logs/`, `checkpoints/` et `output/` sont créés automatiquement s'ils n'existent pas.

## Installation locale

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Variables d'environnement

Créer un fichier `.env` à partir de `.env.example` ou exporter les variables directement :

```bash
export SCRAPER_DATA_DIR=runtime_data
export EMAIL_USER=example@gmail.com
export EMAIL_PASSWORD=gmail_app_password
export EMAIL_TO=receiver@gmail.com
```

Notes Gmail :
- utiliser un mot de passe d'application Gmail
- ne jamais hardcoder les identifiants dans le code

## Exécution

Commande officielle :

```bash
python avibase_exact_excel_scraper_FR_OISEAUX.py --output output/avibase_maroc_exact.xlsx
```

Exemple de test rapide :

```bash
python avibase_exact_excel_scraper_FR_OISEAUX.py --output output/test.xlsx --limit 10
```

Options utiles :
- `--limit 10` pour un test sur un sous-ensemble
- `--save-every 5` pour définir la fréquence de sauvegarde partielle
- `--checkpoint checkpoints/progress.json` pour changer le chemin du checkpoint

## Persistance sur Railway

Sur Railway, le système de fichiers est éphémère par défaut. Pour conserver `logs/`, `checkpoints/` et `output/` entre redémarrages et redéploiements, il faut attacher un Volume au service.

Le scraper détecte automatiquement `RAILWAY_VOLUME_MOUNT_PATH` si un volume est attaché. Il est donc recommandé de monter le volume sur :
- `/app/runtime_data`

Avec cette configuration :
- `logs/` devient `/app/runtime_data/logs`
- `checkpoints/` devient `/app/runtime_data/checkpoints`
- `output/` devient `/app/runtime_data/output`

En local, `SCRAPER_DATA_DIR=runtime_data` reste pratique pour écrire dans un dossier dédié du projet.

## Reprise automatique

Le scraper sauvegarde sa progression dans `checkpoints/progress.json`.

Le checkpoint contient notamment :
- les `avibase_id` déjà traités
- la dernière espèce traitée
- les statistiques de progression
- les données collectées jusqu'à présent
- les chemins de sortie

Au redémarrage :
- le checkpoint est détecté automatiquement
- les espèces déjà traitées ne sont pas retraitées
- la progression restante est recalculée
- un export partiel peut déjà être disponible dans `output/`

## Logs

Le projet écrit des logs :
- en console
- dans `logs/scraper.log`

Les logs couvrent :
- démarrage du scraper
- nombre total d'espèces
- progression espèce par espèce
- retries HTTP et timeouts
- erreurs partielles
- sauvegardes checkpoint
- exports Excel partiels et finaux
- durée totale
- erreurs critiques avec traceback

## Checkpoints et sauvegardes

Fichiers attendus en cours d'exécution :
- `checkpoints/progress.json`
- `output/avibase_maroc_exact_partial.xlsx`
- `output/avibase_maroc_exact.xlsx`
- `logs/scraper.log`

Le checkpoint JSON est sauvegardé à chaque espèce traitée.

L'Excel partiel est sauvegardé périodiquement, par défaut toutes les 5 espèces.

## Notifications email

Le module `utils/mailer.py` envoie un email :
- à la fin du scraping en cas de succès
- en cas d'erreur critique

Le mail contient :
- le statut
- la durée
- le nombre d'espèces traitées
- le fichier généré
- le nombre d'erreurs
- le traceback complet en cas de crash

Si possible, le fichier Excel final ou partiel est joint.

## Déploiement Railway

Étapes exactes :

1. Push le projet sur GitHub.
2. Ouvrir Railway et créer un nouveau projet.
3. Cliquer `New` puis `GitHub Repo`.
4. Sélectionner le repository GitHub privé.
5. Railway détecte le dépôt et utilise `railway.json`.
6. Vérifier ou renseigner les paramètres suivants :

```bash
python avibase_exact_excel_scraper_FR_OISEAUX.py --output output/avibase_maroc_exact.xlsx
```

7. Ajouter un volume au service et le monter sur :
- `/app/runtime_data`

8. Ajouter les variables d'environnement :
- `EMAIL_USER`
- `EMAIL_PASSWORD`
- `EMAIL_TO`

9. Déployer le service.

Le fichier `railway.json` fournit :
- le `startCommand`
- la restart policy `ON_FAILURE`
- jusqu'à 5 retries automatiques en cas de crash
- un `drainingSeconds` de 300 secondes pour laisser au scraper le temps de sauvegarder son checkpoint lors d'un arrêt

## Étapes Railway détaillées

1. Aller sur `https://railway.com`.
2. Se connecter avec GitHub.
3. Créer un projet vide ou cliquer `New Project`.
4. Choisir `GitHub Repo`.
5. Si le repo privé n'apparaît pas, autoriser l'app GitHub Railway sur le repo.
6. Sélectionner `anassbougaiouar/scrap_avibase_oiseau`.
7. Ouvrir le service créé, puis l'onglet `Variables`.
8. Ajouter :
- `EMAIL_USER`
- `EMAIL_PASSWORD`
- `EMAIL_TO`
9. Ouvrir le canvas du projet, puis ajouter un `Volume`.
10. Attacher ce volume au service du scraper.
11. Définir le mount path : `/app/runtime_data`
12. Lancer ou relancer le déploiement.
13. Vérifier les logs du service.

## GitHub autodeploy

Railway déploie automatiquement les nouveaux commits quand le service est connecté à un repo GitHub. Si l'autodeploy ne fonctionne pas, vérifier :
- qu'au moins un membre du projet Railway a un accès contributeur au repo
- que l'app GitHub Railway a accès au repo privé
- qu'aucune mise à jour de permissions GitHub n'est en attente

Le commit poussé sur `main` déclenchera alors un nouveau déploiement automatiquement.

## GitHub

Exemple de séquence si le dossier n'est pas encore un dépôt Git :

```bash
git init
git add .
git commit -m "Prepare Avibase scraper for Railway deployment"
git branch -M main
git remote add origin <URL_DU_REPO_GITHUB>
git push -u origin main
```

## Compatibilité Railway

Le scraper est compatible Railway :
- sans navigateur
- sans Selenium
- sans Playwright
- avec `requests` + `BeautifulSoup`
- avec volume persistant pour les checkpoints et exports

## Limites connues

- Le scraping dépend de la structure HTML actuelle d'Avibase et d'Oiseaux.net.
- Les emails Gmail nécessitent un mot de passe d'application valide.
- Le build de l'index Oiseaux.net est refait à chaque redémarrage.
- Un arrêt brutal pendant l'écriture d'un export partiel peut laisser uniquement le checkpoint JSON comme source de reprise.
- Le volume Railway doit être configuré manuellement depuis le dashboard, car `railway.json` ne crée pas le volume à lui seul.

## Pistes d'optimisation futures

- Mettre en cache l'index Oiseaux.net entre les redémarrages
- Ajouter des tests unitaires ciblés sur les parseurs HTML
- Ajouter une compression ou une rotation avancée des checkpoints
- Produire des exports CSV intermédiaires par feuille si besoin d'audit
