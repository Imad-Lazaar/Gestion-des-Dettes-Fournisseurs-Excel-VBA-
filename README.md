# Gestion des dettes fournisseurs (Excel + VBA)

Application Excel pour suivre les factures fournisseurs, les règlements et les retards de paiement, avec un tableau de bord interactif et des macros VBA.

## Aperçu

<img width="856" height="488" alt="TB" src="https://github.com/user-attachments/assets/ed9dea0c-557c-4c72-aabe-db54a0c379b5" />

*Tableau de bord : total achats, total payé, reste à payer, retards, factures échues, soldes par fournisseur et par catégorie. Des segments filtrent par fournisseur et par mois.*

<img width="1267" height="380" alt="Liste des dettes" src="https://github.com/user-attachments/assets/f857dab5-1b98-48b9-8f96-c6273bab0da0" />
*Feuille Dette : formulaire de saisie et tableau des factures avec montant payé, solde, échéance, retard et statut calculés automatiquement.*

<img width="903" height="475" alt="Réglement" src="https://github.com/user-attachments/assets/c98146e6-2101-4556-a1d0-e7c7d9949821" />

*Feuille Règlement : le numéro de facture remplit le fournisseur et le solde ; chaque paiement est enregistré avec sa référence et son mode.*

<img width="1215" height="553" alt="Liste Des Fournisseurs" src="https://github.com/user-attachments/assets/fa98ca53-eeec-4a8e-992c-23a7ca50ae77" />

*Feuille Paramètres : liste des fournisseurs (avec leur catégorie) et date de référence.*

## Fonctionnement

- **Dette** : on saisit une facture (numéro, date, fournisseur en liste déroulante, délai, montant). La catégorie est reprise automatiquement des paramètres.
- **Règlement** : on saisit le numéro de facture, et le fournisseur et le solde s'affichent. On enregistre la date, la référence, le montant et le mode de paiement.
- **Calculs automatiques** dans le tableau des dettes :
  - Montant payé = somme des règlements de la facture ;
  - Solde = montant à payer − montant payé ;
  - Échéance = date de facture + délai ;
  - Retard = jours écoulés depuis l'échéance, tant qu'il reste un solde ;
  - Statut : *Payé*, *Non échue*, *Échue* ou *Trop-perçu*.
- **Date de référence** : la date du jour par défaut (`AUJOURDHUI()`), qu'on peut remplacer par une date fixe pour produire un état à une date donnée.
- **Tableau de bord** : alimenté par des tableaux croisés dynamiques (feuille *Traitement*), avec un bouton d'actualisation.
- **Macros VBA** : ajout d'une dette et d'un règlement, annulation de la saisie, navigation entre les feuilles, actualisation.
  - L'ajout d'une dette vérifie les champs vides et refuse un numéro de facture déjà utilisé.

## Ce que j'ai modifié par rapport au tutoriel

- Reconstruction complète avec des données marocaines (fournisseurs, montants en DH).
- Une seule date de référence pour le retard et le statut.
- Statut « Trop-perçu » pour les règlements supérieurs au montant dû, et retard jamais négatif.
- Contrôle des doublons de numéro de facture à la saisie.
- Suppression du calcul du DPO, que j'ai jugé trop lourd pour l'intérêt qu'il apportait.
- Nettoyage des liens externes et des macros inutilisées.

## Utilisation

1. Ouvrir le fichier `.xlsm` et **activer les macros**.
2. Saisir les factures dans *Dette* et les paiements dans *Règlement*.
3. Cliquer sur le bouton d'actualisation du tableau de bord.
