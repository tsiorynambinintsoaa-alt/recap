# Générateur Multi-Clients - Rapports de Production

Outil d'automatisation pour générer les récapitulatifs de production par client à partir du fichier SUIVI.

## Fonctionnalités principales

- Traitement de plusieurs clients en une seule fois
- Génération d'un fichier Excel bien structuré et mis en forme automatiquement
- Création d'un fichier HTML avec des graphiques interactifs (barres empilées, KPIs, mode sombre/clair)
- Interface graphique simple et intuitive
- Journal d'activité en temps réel

## Feuilles Excel générées

- RECAP GLOBAL (avec total général)
- A et ST
- RAMASSAGE PAR JOUR / SEMAINE / MOIS
- RAMASSAGE PAR MODELES
- DISPATCH
- RAMASSAGE PAR PRESTATAIRE
- STATS_TEMPS
- TIMING PAR MODELE

## Prérequis

- Python 3.8 ou supérieur
- Bibliothèques : pandas, openpyxl, ttkbootstrap

## Installation

1. Installer les dépendances :

```bash
pip install pandas openpyxl ttkbootstrap
