# AutoRCI_Control

AutoRCI_Control est une solution Python d'automatisation du rapprochement entre les donnees ERP Microsoft Navision/Incadea et les flux RCI Banque.

Le projet detecte les fichiers disponibles, extrait les factures et avoirs depuis l'ERP, les fichiers RCI et les PDF banque, puis genere un rapport de controle exploitable par l'equipe facturation.

## Objectif Metier

L'objectif est de fiabiliser le controle quotidien des factures de pieces de rechange financees ou transmises a RCI Banque.

Le pipeline permet de :

- comparer les factures ERP avec les donnees RCI et PDF ;
- identifier les factures manquantes cote RCI ;
- detecter les anomalies de montant, de date ou de duplication ;
- produire un rapport Excel professionnel ;
- alimenter un historique Power BI ;
- preparer l'envoi automatique du rapport par email ;
- archiver les fichiers traites en execution reelle.

## Architecture Des Dossiers

```text
AutoRCI_Control/
+-- input/
|   +-- erp/
|   +-- rci/
|   +-- pdf/
+-- samples/
|   +-- erp/
|   +-- rci/
|   +-- pdf/
+-- reference/
+-- output/
|   +-- reports/
|   +-- powerbi/
|   +-- anomalies/
+-- archive/
|   +-- erp/
|   +-- rci/
|   +-- pdf/
+-- logs/
+-- src/
+-- webapp/
+   +-- app.py
+   +-- templates/
+   +-- static/
+-- tests/
+-- config/
|   +-- config.yaml
+-- requirements.txt
+-- README.md
```

## Important

Les fichiers places dans `samples/` sont uniquement des exemples de format.

En production, les nouveaux fichiers quotidiens doivent etre deposes dans :

- `input/erp`
- `input/rci`
- `input/pdf`

Aucun nom de fichier n'est code en dur. Le systeme detecte automatiquement les fichiers presents dans les dossiers.

## Samples Et Input

`samples/` sert uniquement au developpement, aux tests et aux validations de format. Les fichiers qui s'y trouvent ne sont jamais archives par le pipeline.

`input/` est le dossier de production. Chaque jour, les nouveaux exports ERP, fichiers RCI et PDF doivent y etre deposes avant execution.

En execution reelle, les fichiers traites depuis `input/` peuvent etre deplaces vers `archive/` apres succes, sauf si l'option `--dry-run` est utilisee.

## Installation

Depuis le dossier du projet :

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Les dependances principales sont `pandas`, `openpyxl` et `pdfplumber`.

## Dashboard Flask

Une application web interne Flask + Bootstrap est disponible pour utiliser AutoRCI sans passer par le terminal.

Elle permet de :

- consulter les KPI du dernier traitement ;
- deposer les fichiers ERP, RCI, PDF et le referentiel RCI ;
- lancer une reconciliation en `dry-run` ou en execution reelle ;
- filtrer les resultats et les ecarts ;
- telecharger les rapports, exports Power BI, anomalies et logs ;
- lire le dernier log d'execution.

Installation :

```powershell
pip install -r requirements.txt
```

Lancement :

```powershell
python webapp/app.py
```

Ou :

```powershell
flask --app webapp/app run
```

URL :

```text
http://127.0.0.1:5000
```

Par defaut, le formulaire de traitement coche `dry-run` afin d'eviter l'envoi email et l'archivage accidentels.

## Execution En Mode Test

Pour tester le pipeline avec les fichiers d'exemple sans email et sans archivage :

```powershell
python src/main.py --use-samples --dry-run
```

Ce mode lit les fichiers dans `samples/erp`, `samples/rci` et `samples/pdf`.

## Execution Reelle

Apres depot des fichiers quotidiens dans `input/` :

```powershell
python src/main.py
```

Variantes utiles :

```powershell
python src/main.py --dry-run
```

`--dry-run` execute les controles et genere les sorties, mais n'envoie pas d'email et n'archive pas les fichiers.

Options de diagnostic utiles :

```powershell
python src/main.py --use-samples --dry-run --ignore-pdf
python src/main.py --use-samples --dry-run --ignore-pdf --date-from 2026-04-29 --date-to 2026-05-05
python src/main.py --use-samples --dry-run --ignore-pdf --no-date-filter
python src/main.py --use-samples --dry-run --debug-invoice VF384312
python src/main.py --use-samples --dry-run --debug-reference
```

`--ignore-pdf` isole le rapprochement ERP vs RCI TXT/CSV/Excel.  
`--date-from` et `--date-to` forcent manuellement la periode de rapprochement.  
`--no-date-filter` desactive le filtre de periode et compare tout l'ERP extrait.  
`--debug-invoice` ecrit dans les logs toutes les lignes trouvees pour une facture cote ERP, RCI TXT, PDF, RCI consolide, ainsi que la raison du statut final.
`--debug-reference` inspecte le referentiel RCI et genere un CSV `output/anomalies/reference_debug_YYYYMMDD_HHMMSS.csv`.

## Rapport Excel Genere

Le rapport Excel est genere dans :

```text
output/reports/
```

Nom du fichier :

```text
Rapport_Reconciliation_RCI_YYYY-MM-DD_HHMM.xlsx
```

Le classeur contient :

- `Synthese` : indicateurs globaux du traitement ;
- `Detail rapprochement` : toutes les factures analysees ;
- `Factures manquantes RCI` : factures presentes ERP mais absentes RCI/PDF ;
- `Anomalies` : anomalies de montant ou de date ;
- `Doublons` : factures detectees plusieurs fois ;
- `RCI seulement` : factures presentes cote RCI/PDF mais absentes ERP ;
- `Synthese par concessionnaire` : vision par client/concessionnaire.

Les statuts sont colores :

- `OK` : vert ;
- `MANQUANTE_RCI` : orange ;
- `ANOMALIE_MONTANT` et `ANOMALIE_DATE` : rouge ;
- `DOUBLON` : violet ;
- `RCI_SEULEMENT` : bleu.

## Fichier Power BI

Le fichier historique Power BI est :

```text
output/powerbi/reconciliation_history.csv
```

Il est mis a jour a chaque execution. Les nouvelles lignes sont ajoutees sans supprimer l'historique.

Colonnes principales :

- `processing_date`
- `processing_run_id`
- `invoice_number`
- `document_type`
- `customer_code`
- `customer_name`
- `amount_erp`
- `amount_rci`
- `amount_pdf`
- `amount_gap`
- `erp_date`
- `pdf_invoice_date`
- `due_date`
- `origin`
- `status`
- `priority`
- `action_recommandee`

Ce fichier permet de construire un dashboard Power BI avec :

- montant total controle ;
- montant manquant RCI ;
- nombre d'anomalies ;
- taux de rapprochement ;
- evolution par jour ;
- top concessionnaires avec ecarts.

## Configuration Email

La configuration email se trouve dans :

```text
config/config.yaml
```

Exemple :

```yaml
email:
  enabled: true
  sender: "autorcicontrol@entreprise.com"
  recipients:
    - "facturation@entreprise.com"
  smtp_host: "smtp.entreprise.com"
  smtp_port: 587
  smtp_username: "autorcicontrol"
  smtp_password_env_var: "AUTORCI_SMTP_PASSWORD"
  use_tls: true
  use_ssl: false
```

Le mot de passe SMTP ne doit jamais etre ecrit dans le code ni dans le depot.

Definir le mot de passe dans une variable d'environnement Windows :

```powershell
setx AUTORCI_SMTP_PASSWORD "mot_de_passe_smtp"
```

En mode `--dry-run`, aucun email n'est envoye. Le sujet, le corps et la piece jointe prevus sont seulement affiches dans les logs.

## Planification Windows Task Scheduler

Pour executer le traitement automatiquement chaque jour a 21h30 :

1. Ouvrir Windows Task Scheduler.
2. Creer une nouvelle tache.
3. Onglet `Triggers` : ajouter un declencheur quotidien a `21:30`.
4. Onglet `Actions` : ajouter une action `Start a program`.

Exemple de configuration :

Program/script :

```text
C:\chemin\vers\AutoRCI_Control\.venv\Scripts\python.exe
```

Arguments :

```text
C:\chemin\vers\AutoRCI_Control\src\main.py
```

Start in :

```text
C:\chemin\vers\AutoRCI_Control
```

Pour tester la tache planifiee sans email ni archivage, utiliser temporairement :

```text
C:\chemin\vers\AutoRCI_Control\src\main.py --dry-run
```

## Gestion Des Logs

Les logs sont generes dans :

```text
logs/
```

Chaque execution cree un fichier :

```text
autorcicontrol_YYYYMMDD_HHMMSS.log
```

Les logs indiquent :

- les fichiers detectes ;
- les lignes extraites par source ;
- les statuts de rapprochement ;
- les chemins des rapports generes ;
- le statut email ;
- l'archivage ;
- la synthese finale.

## Limites Et Ameliorations Futures

Limites actuelles :

- les formats ERP/RCI/PDF tres differents peuvent necessiter l'ajout de nouvelles regles de detection ;
- le rapprochement se base principalement sur `invoice_number` normalise ;
- la performance depend du volume PDF et du temps d'extraction `pdfplumber` ;
- l'envoi email necessite une configuration SMTP valide.

Ameliorations possibles :

- enrichir les controles par concessionnaire, VIN/chassis ou contrat ;
- ajouter un tableau de bord Power BI preconfigure ;
- ajouter une interface graphique de suivi ;
- ajouter une rotation automatique des logs ;
- ajouter une notification Teams/Outlook en complement de l'email ;
- ajouter des tests d'integration avec des jeux de donnees anonymises de production.
