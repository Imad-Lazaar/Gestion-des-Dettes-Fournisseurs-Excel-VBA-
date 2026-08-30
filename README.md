# Gestion des Dettes Fournisseurs (Excel/VBA)

Outil Excel pour suivre les factures fournisseurs, les paiements, et le retard de paiement (aging), avec un tableau de bord.

## Fonctionnalités

- Suivi des factures fournisseurs et des paiements
- Calcul automatique du solde restant à payer
- Classement des factures par retard (échue, non échue, 30/60/90 jours)
- Calcul du **DPO** (délai moyen de paiement)
- Tableau de bord avec graphiques et filtres
- Boutons (macros VBA) pour ajouter une facture/un paiement en un clic
- ---
## Tableau de bord
Vue d'ensemble : total des achats, montant payé, factures à venir, retards, taux de paiement, balance âgée, évolution du DPO, top fournisseurs impayés, et solde par catégorie.

<img width="1287" height="486" alt="Tableau de bord" src="https://github.com/user-attachments/assets/8c6fc839-311b-44d3-bea3-9b14e81cd889" />

---
## Liste de dettes
Formulaire pour ajouter une nouvelle facture fournisseur, avec la liste complète des factures et leur statut (payé, échue, en retard).

<img width="859" height="601" alt="Liste de réglements" src="https://github.com/user-attachments/assets/71c3b103-1184-41f2-8284-1f8f6cfc7c32" />

---
## Liste des règlements
Formulaire pour enregistrer un paiement, avec l'historique de tous les règlements effectués (référence, montant, mode de paiement).

<img width="1232" height="568" alt="Liste de dettes" src="https://github.com/user-attachments/assets/e1bcb4c6-85f6-460b-bc5a-50f8c5ddf735" />

---
## Liste de fournisseurs
Répertoire des fournisseurs (nom, catégorie, coordonnées) et la date de référence utilisée pour les calculs.

<img width="1238" height="506" alt="Liste de fournisseurs" src="https://github.com/user-attachments/assets/2ad66caf-8301-48c9-a8b3-1684d8631ce7" />

---

## Compétences utilisées

- **Excel avancé** : tableaux structurés, formules (SUMIF, IF imbriqués), liste déroulante dynamique
- **Finance** : logique d'aging, calcul du DPO
- **Tableaux croisés dynamiques & graphiques** : 9 TCD, 5 graphiques, un slicer
- **VBA** : macros pour automatiser la saisie et la navigation
