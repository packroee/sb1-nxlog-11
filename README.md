# Analyseur de Configuration NXLog

Ce script Python analyse les fichiers de configuration nxlog et affiche les paramètres sous forme de tableau avec descriptions. Il inclut également une fonctionnalité de cartographie des flux pour visualiser les routes et connexions entre les sections.

## Installation

```bash
python3 -m pip install -r requirements.txt
```

Ou installer directement la dépendance principale :

```bash
python3 -m pip install tabulate
```

**Note**: Si le module `tabulate` n'est pas disponible, le script fonctionnera quand même avec un format de tableau simplifié ou les formats CSV/JSON.

## Utilisation

### Analyser un fichier de configuration existant

```bash
python3 nxlog_analyzer.py /path/to/nxlog.conf
```

### Créer un fichier d'exemple

```bash
python3 nxlog_analyzer.py --create-sample
```

### Afficher les statistiques

```bash
python3 nxlog_analyzer.py nxlog.conf --stats
```

### Afficher la cartographie des flux

```bash
python3 nxlog_analyzer.py nxlog.conf --flows
```

### Analyser un répertoire complet

```bash
# Analyser tous les fichiers .conf dans le répertoire data
python3 nxlog_analyzer.py --directory data

# Avec statistiques et cartographie des flux
python3 nxlog_analyzer.py --directory data --stats --flows
```

### Formats de sortie

- **Table** (par défaut): Affichage tabulaire
```bash
python3 nxlog_analyzer.py nxlog.conf --format table
```

- **CSV**: Format CSV
```bash
python3 nxlog_analyzer.py nxlog.conf --format csv
```

- **JSON**: Format JSON
```bash
python3 nxlog_analyzer.py nxlog.conf --format json
```

### Génération de rapports

```bash
# Rapport Excel complet avec cartographie des flux
python3 nxlog_analyzer.py --directory data --excel-file rapport_complet.xlsx

# Fichiers CSV séparés + cartographie des flux
python3 nxlog_analyzer.py --directory data --csv-multiple --flows-csv

# Génération de diagrammes Graphviz
### Génération de diagrammes visuels

```bash
# Générer les fichiers Graphviz
python3 nxlog_analyzer.py --directory data --graphviz

# Puis générer les images (nécessite Graphviz installé)
cd output
./nxlog_sample_generate_images.sh
```

python3 nxlog_analyzer.py --directory data --graphviz
```

## Fonctionnalités

- ✅ Parse les fichiers de configuration nxlog
- ✅ Affiche les paramètres par section
- ✅ Descriptions détaillées des paramètres
- ✅ Statistiques de configuration
- ✅ Formats de sortie multiples (table, CSV, JSON)
- ✅ Gestion des commentaires
- ✅ Création d'exemples de configuration
- ✅ Support des modules principaux (Input, Output, Route, Extension, Processor)
- ✅ Fonctionne avec ou sans le module `tabulate`
- ✅ **Cartographie des flux et des routes**
- ✅ **Analyse des connexions entre sections**
- ✅ **Traitement de répertoires multiples**
- ✅ **Rapports Excel avec onglets séparés**
- ✅ **Export CSV multiple avec cartographie des flux**
- ✅ **Génération de diagrammes Graphviz pour visualisation**
- ✅ **Scripts automatiques de génération d'images**
- ✅ **Cartographie de synthèse globale combinant tous les fichiers**

## Exemples de sortie

### Format tableau (avec tabulate)
```
+----------+-------------+---------------+------------------+--------------------------------+
| Section  | Nom Section | Paramètre     | Valeur           | Description                    |
+==========+=============+===============+==================+================================+
| Input    | eventlog    | Module        | im_msvistalog    | Type de module utilisé         |
+----------+-------------+---------------+------------------+--------------------------------+
| Input    | file        | Module        | im_file          | Type de module utilisé         |
+----------+-------------+---------------+------------------+--------------------------------+
```

### Format tableau (sans tabulate)
```
+----------+-------------+---------------+------------------+--------------------------------+
| Section  | Nom Section | Paramètre     | Valeur           | Description                    |
+----------+-------------+---------------+------------------+--------------------------------+
| Input    | eventlog    | Module        | im_msvistalog    | Type de module utilisé         |
| Input    | file        | Module        | im_file          | Type de module utilisé         |
+----------+-------------+---------------+------------------+--------------------------------+
```

### Statistiques
```
==================================================
STATISTIQUES DE CONFIGURATION
==================================================
Nombre total de paramètres: 25
Nombre de sections: 6
Nombre de modules: 5

Sections trouvées: Input, Output, Route, Extension, Processor
Modules utilisés: im_msvistalog, im_file, om_udp, om_file, xm_csv
==================================================
```

### Cartographie des flux
```
================================================================================
CARTOGRAPHIE DES FLUX - NXLOG_SAMPLE
================================================================================
📊 RÉSUMÉ:
  • Routes: 2
  • Sections: 6
  • Flux: 4
  • Inputs: 2
  • Outputs: 2
  • Processors: 1
  • Extensions: 1
  • Sections non connectées: 1

🔄 FLUX DE DONNÉES:
+-------+-----------+-------------+---------------+---+-------------+----------+-------------+----------+-----------+
| Route | Source    | Type Source | Module Source |   | Destination | Type Dest| Module Dest | Priorité | Condition |
+=======+===========+=============+===============+===+=============+==========+=============+==========+===========+
| main  | eventlog  | Input       | im_msvistalog | → | pattern     | Processor| pm_pattern  | 1        | N/A       |
| main  | file      | Input       | im_file       | → | pattern     | Processor| pm_pattern  | 1        | N/A       |
| main  | pattern   | Processor   | pm_pattern    | → | syslog      | Output   | om_udp      | 1        | N/A       |
| main  | pattern   | Processor   | pm_pattern    | → | fileout     | Output   | om_file     | 1        | N/A       |
+-------+-----------+-------------+---------------+---+-------------+----------+-------------+----------+-----------+
```

### Diagrammes Graphviz
```bash
# Génère des fichiers .dot et scripts de génération d'images
python3 nxlog_analyzer.py --directory data --graphviz

# Fichiers générés:
# - nxlog_sample_flow.dot (définition du graphique)
# - nxlog_sample_generate_images.sh (script de génération)
# - nxlog_synthesis_flow.dot (cartographie de synthèse globale)
# - nxlog_synthesis_generate_images.sh (script pour la synthèse)
```

### Cartographie de Synthèse
```bash
# La cartographie de synthèse combine automatiquement tous les fichiers
python3 nxlog_analyzer.py --directory data --graphviz

# Génération des images de synthèse
cd output
./nxlog_synthesis_generate_images.sh

# Fichiers de synthèse générés:
# - nxlog_synthesis_flow.png (image bitmap)
# - nxlog_synthesis_flow.svg (image vectorielle)
# - nxlog_synthesis_flow.pdf (document imprimable)
```
### Rapports Excel
- **Onglets de configuration** : Un onglet par fichier `.conf` avec tous les paramètres
- **Onglets de flux** : Cartographie des flux pour chaque fichier (`_Flux`, `_Sections`)
- **Onglet statistiques** : Vue consolidée avec statistiques globales et par fichier

## Paramètres supportés

Le script reconnaît plus de 100 paramètres nxlog courants avec leurs descriptions, incluant:

- **Modules**: im_file, om_file, im_msvistalog, om_udp, xm_csv, etc.
- **Fichiers**: File, LogFile, CertFile, etc.
- **Réseau**: Host, Port, Protocol, SSL, etc.
- **Formats**: InputType, OutputType, Format, etc.
- **Performance**: BufferSize, FlushInterval, PollInterval, etc.
- **Sécurité**: SSL, CertFile, AllowUntrusted, etc.

## Structure du fichier de configuration

Le script analyse les sections suivantes:
- `<Input>` - Modules d'entrée
- `<Output>` - Modules de sortie  
- `<Route>` - Routes de traitement
- `<Extension>` - Extensions
- `<Processor>` - Processeurs
- Paramètres globaux

## Dépendances

## Cartographie des flux

La fonctionnalité de cartographie des flux analyse automatiquement :
- **Routes** : Chemins de traitement définis dans les sections `<Route>`
- **Connexions** : Relations entre Input → Processor → Output
- **Priorités** : Ordre d'exécution des routes
- **Conditions** : Conditions d'activation des routes
- **Sections non connectées** : Sections définies mais non utilisées dans les routes

### Format des routes supporté
- `input1, input2 => output1, output2`
- `input1 => processor1 => output1`
- `input1, input2 => processor1, processor2 => output1, output2`

## Visualisation Graphviz

Le script génère des diagrammes visuels au format Graphviz (.dot) qui peuvent être convertis en images :

### Types de diagrammes générés

1. **Diagrammes individuels** : Un diagramme par fichier de configuration
   - Nom : `{nom_fichier}_flow.dot`
   - Contenu : Flux spécifiques à ce fichier
   - Script : `{nom_fichier}_generate_images.sh`

2. **Cartographie de synthèse** : Vue d'ensemble globale
   - Nom : `nxlog_synthesis_flow.dot`
   - Contenu : Tous les flux de tous les fichiers combinés
   - Script : `nxlog_synthesis_generate_images.sh`
   - Caractéristiques :
     - Sous-graphes colorés par fichier de configuration
     - Sections préfixées par le nom du fichier
     - Connexions inter-fichiers visibles
     - Statistiques globales intégrées

### Installation de Graphviz
```bash
# Ubuntu/Debian
sudo apt-get install graphviz

# CentOS/RHEL
sudo yum install graphviz

# macOS
brew install graphviz

# Windows
# Télécharger depuis https://graphviz.org/download/
```

### Formats de sortie supportés
- **PNG** : Images bitmap haute qualité
- **SVG** : Images vectorielles (redimensionnables)
- **PDF** : Documents imprimables
- **DOT** : Format source Graphviz

### Caractéristiques des diagrammes
- **Couleurs par type** : Input (vert), Output (rouge), Processor (bleu), Extension (jaune)
- **Sections non connectées** : Affichées en pointillés
- **Routes colorées** : Chaque route a sa propre couleur
- **Informations détaillées** : Priorités et conditions sur les connexions
- **Légende intégrée** : Explication des couleurs et symboles
- **Cartographie de synthèse** : Vue globale avec clusters par fichier de configuration
- **Identification claire** : Sections préfixées par le nom du fichier source
- **Couleurs de fond** : Chaque fichier a sa propre couleur de cluster
- **Tooltips informatifs** : Informations détaillées au survol (formats supportés)

## Dépendances

- Python 3.6+
- tabulate (optionnel, pour un meilleur affichage des tableaux)
- openpyxl (optionnel, pour la génération de fichiers Excel)
- graphviz (optionnel, pour la génération d'images à partir des fichiers .dot)

## Compatibilité

Le script fonctionne dans différents environnements :
- Systèmes avec `pip` disponible
- Systèmes avec seulement `python3 -m pip`
- Environnements sans `tabulate` (utilise un format de tableau simplifié)

## Workflow recommandé pour l'analyse complète

```bash
# 1. Analyse complète avec tous les formats de sortie
python3 nxlog_analyzer.py --directory data --stats --flows --graphviz --excel rapport_complet.xlsx

# 2. Génération des images individuelles
cd output
for script in *_generate_images.sh; do
    ./"$script"
done

# 3. Génération de l'image de synthèse
./nxlog_synthesis_generate_images.sh

# 4. Consultation des résultats
# - rapport_complet.xlsx : Analyse détaillée dans Excel
# - *.png/*.svg/*.pdf : Diagrammes visuels
# - nxlog_synthesis_flow.* : Vue d'ensemble globale
```

## Cas d'usage de la cartographie de synthèse

La cartographie de synthèse est particulièrement utile pour :

- **Architectures multi-fichiers** : Visualiser les interactions entre plusieurs configurations
- **Audit de sécurité** : Identifier les flux de données à travers l'infrastructure
- **Documentation système** : Créer une vue d'ensemble pour la documentation
- **Troubleshooting** : Comprendre rapidement les chemins de données complexes
- **Migration/Refactoring** : Analyser l'impact des changements sur l'ensemble du système
- **Formation** : Expliquer l'architecture globale aux nouveaux membres de l'équipe