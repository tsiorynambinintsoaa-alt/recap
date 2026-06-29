# 🧵 Générateur de Rapports de Production – Artisanat Raphia

Cet outil automatise la génération de récapitulatifs de production par client à partir d'un fichier Excel de suivi. Il est spécialement conçu pour le suivi de la production artisanale (articles en raphia), permettant d'exporter les données sous formes de fichiers **Excel mis en forme** et de **rapports HTML interactifs**.

---

## 🚀 Fonctionnalités principales

- **Traitement de masse :** Gestion de plusieurs clients simultanément en une seule exécution.
- **Exports Excel automatisés :** Génération de classeurs structurés, nettoyés et mis en forme de manière professionnelle.
- **Tableaux de bord HTML interactifs :** Visualisation dynamique avec graphiques (barres empilées), indicateurs clés de performance (KPIs) et bascule entre mode sombre et mode clair.
- **Interface Graphique (GUI) :** Prise en main simple, intuitive et accessible à tous les utilisateurs.
- **Suivi en direct :** Journal d'activité (logs) intégré pour suivre l'avancement des tâches en temps réel.

---

# 📸 Aperçu

## 📊 Rapport Excel généré

<p align="center">
  <img src="images/fichier_excel.png" alt="Rapport Excel" width="100%">
</p>

Le fichier Excel contient un récapitulatif complet de la production avec plusieurs feuilles d'analyse, des tableaux formatés automatiquement et des indicateurs de suivi.

---

## 🌐 Tableau de bord HTML interactif

<p align="center">
  <img src="images/fichier_html.png" alt="Dashboard HTML" width="100%">
</p>

Le tableau de bord HTML permet de visualiser les indicateurs clés de production grâce à des graphiques interactifs, des cartes KPI et une interface moderne avec mode clair/sombre.

---

## 📊 Structure des Feuilles Excel Générées

L'outil génère un fichier complet pour chaque client, comprenant les onglets suivants :

| Nom de la feuille | Description |
| :--- | :--- |
| **RECAP GLOBAL** | Vue d'ensemble avec totaux généraux : quantités commandées, dispatchées, livrées, reste à dispatcher et reste à livrer. |
| **A et ST** | Analyse fine de la production répartie entre l'Atelier principal et la Sous-Traitance. |
| **RAMASSAGE TEMP** | Suivi chronologique des ramassages triés par jour, semaine et mois. |
| **RAMASSAGE MODELES** | Analyse quotidienne des volumes de ramassage détaillés par modèle de produit. |
| **DISPATCH** | Suivi de la distribution et de l'affectation des pièces. |
| **RAMASSAGE PRESTATAIRE** | Suivi individuel par artisan : pièces livrées et reste à livrer. |
| **TIMING PAR MODELE** | Analyse de la performance et du temps de fabrication par pièce (moyenne et médiane). |

---

## 🛠️ Prérequis

- **Python :** Version 3.8 ou supérieure.
- **Bibliothèques requises :** `pandas`, `openpyxl`, `ttkbootstrap`

---

## 📦 Installation & Configuration

1. Téléchargez ou clonez le projet.
2. Installez les dépendances :

```bash
pip install pandas openpyxl ttkbootstrap
