# 🧵 Générateur de Rapports de Production – Artisanat Raphia

Cet outil automatise la génération de récapitulatifs de production par client à partir d'un fichier Excel de suivi. Il est spécialement conçu pour le suivi de la production artisanale (articles en raphia), permettant d'exporter les données sous formes de fichiers **Excel mis en forme** et de **rapports HTML interactifs**.

---

## 🚀 Fonctionnalités principales

* **Traitement de masse :** Gestion de plusieurs clients simultanément en une seule exécution.
* **Exports Excel automatisés :** Génération de classeurs structurés, nettoyés et mis en forme de manière professionnelle.
* **Tableaux de bord HTML interactifs :** Visualisation dynamique avec graphiques (barres empilées), indicateurs clés de performance (KPIs) et bascule entre mode sombre et mode clair.
* **Interface Graphique (GUI) :** Prise en main simple, intuitive et accessible à tous les utilisateurs.
* **Suivi en direct :** Journal d'activité (logs) intégré pour suivre l'avancement des tâches en temps réel.

---

## 📊 Structure des Feuilles Excel Générées
L'outil génère un fichier complet pour chaque client, comprenant les onglets suivants :

| Nom de la feuille | Description |
| :--- | :--- |
| **RECAP GLOBAL** | Vue d'ensemble avec totaux généraux : quantités commandées, dispatchées, livrées, reste à dispatcher et reste à livrer. |
| **A et ST** | Analyse fine de la production répartie entre l'**A**telier principal et la **S**ous-**T**raitance. |
| **RAMASSAGE TEMP** | Suivi chronologique des ramassages triés par jour, semaine et mois (ventilé par chef de table en atelier et par sous-traitant). |
| **RAMASSAGE MODELES**| Analyse quotidienne des volumes de ramassage détaillés par modèle de produit. |
| **DISPATCH** | Suivi de la distribution et de l'affectation des pièces. |
| **RAMASSAGE PRESTATAIRE**| Suivi individuel par artisan : pièces livrées et solde restant à dispatcher. |
| **TIMING PAR MODELE** | Analyse de la performance et du temps de fabrication par pièce (calcul des moyennes et médianes). |

---

## 🛠️ Prérequis

* **Python :** Version 3.8 ou supérieure.
* **Bibliothèques requises :** `pandas`, `openpyxl`, `ttkbootstrap`

---

## 📦 Installation & Configuration

1. Téléchargez ou clonez les fichiers du projet.
2. Installez les dépendances nécessaires via votre terminal :

```bash
pip install pandas openpyxl ttkbootstrap
