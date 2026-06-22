# RESERV+ — Guide d'utilisation

## Navigation générale

Après connexion, tous les écrans partagent une **barre latérale sombre** (sidebar)  
avec les modules : Tableau de bord, Réservations, Ressources, Clients, Prestataires,  
Paiements, États & rapports, Paramètres.

En bas de la sidebar : **nom et rôle de l'utilisateur connecté** + bouton de déconnexion.

---

## Module 1 — Tableau de bord

- **4 cartes** : Total réservations / Aujourd'hui / En attente / Montant encaissé
- **Graphique catégories** : répartition Hôtels / Taxis / Vols
- **Dernières réservations** : 5 lignes avec statut coloré
- **Actions rapides** : Nouvelle réservation, Nouveau client, Paiement, Rapports
- Bouton **Actualiser** pour rafraîchir les chiffres

---

## Module 2 — Réservations

### Consulter
- Tableau listant : N°, Client, Période, Montant, Statut (coloré)
- Recherche textuelle (nom client ou numéro)
- Filtres : Toutes | Confirmées | En attente | Annulées

### Actions
| Bouton | Effet |
|--------|-------|
| + Nouvelle réservation | Ouvre le formulaire de création |
| Détail | Affiche le résumé complet |
| Confirmer | Passe de "En attente" à "Confirmée" |
| Annuler | Marque comme "Annulée" (demande confirmation) |
| Paiement | Ouvre le module Paiements |
| Billet | Génère un billet pour la réservation sélectionnée |
| Actualiser | Rafraîchit la liste |

---

## Module 3 — Nouvelle réservation

1. **Sélectionner un client** (liste déroulante)
2. **Choisir la catégorie** (Hôtel / Taxi / Vol)
3. **Choisir la ressource** (filtrée par catégorie, prix auto-récupéré)
4. **Saisir les dates** début et fin
5. **Saisir le nombre de personnes**
6. Cliquer **Vérifier disponibilité**
   - Vert : "Disponible — Aucun conflit"
   - Rouge : "Indisponible — Conflit détecté" (bloque l'enregistrement)
7. Vérifier le **montant calculé automatiquement** (durée × prix × quantité)
8. Cliquer **Enregistrer** → réservation créée + billet généré automatiquement

---

## Module 4 — Ressources

- **Onglets** : Tous / Hôtels / Taxis / Vols
- Chaque ressource affiche : libellé, catégorie, prestataire, prix, capacité, statut
- Bouton **Disponibilité** : bascule entre Disponible / Indisponible
- Bouton **+ Nouvelle ressource** : formulaire de création (prestataire + catégorie requis)

---

## Module 5 — Clients

- Liste avec compteur de réservations
- Recherche par nom, prénom, email, téléphone
- Double-clic : charge la fiche dans le formulaire du bas
- Bouton **+ Nouveau client** : vide la fiche pour saisie
- Bouton **Enregistrer** : crée ou modifie le client
- Bouton **Supprimer** : vérifie l'absence de réservations avant suppression

---

## Module 6 — Prestataires

- Affichage avec type d'activité et compteur de ressources
- Recherche par nom ou type d'activité
- Même logique de fiche que les clients

---

## Module 7 — Paiements

### Vue synthèse
- Total encaissé / Reste à encaisser / Paiements du mois

### Enregistrer un paiement
1. Sélectionner la réservation dans la liste déroulante
2. Le solde restant s'affiche automatiquement
3. Saisir le montant et le mode (Espèces / Carte / Virement / Mobile Money)
4. Cliquer **$ Encaisser**

### Historique
- Liste de tous les paiements avec : réservation, client, montant, mode, solde, état
- Bouton **Générer reçu** : affiche un reçu de paiement formaté

---

## Module 8 — États & rapports

| Rapport | Description |
|---------|-------------|
| Liste réservations | Filtrée par statut et période |
| Chiffre d'affaires | Par catégorie (Hôtel / Taxi / Vol) |
| Taux d'occupation | Par ressource sur une période |
| Journal actions | Admin seulement — traçabilité complète |

Bouton **Exporter** : instructions pour exporter vers Excel.

---

## Module 9 — Paramètres (Admin seulement)

| Paramètre | Description |
|-----------|-------------|
| NOM_ETABLISSEMENT | Affiché sur les billets et reçus |
| DEVISE | GNF (Franc guinéen) par défaut |
| PREFIXE_BILLET | Début du numéro de billet (défaut : BIL) |
| NOM_APPLICATION | Nom affiché dans les formulaires |
| FREQUENCE_SAUVEGARDE | Rappel de sauvegarde (en jours) |

Bouton **Gérer utilisateurs** : ouvre le formulaire de gestion des comptes.

---

## Codes couleurs des statuts

| Statut | Couleur | Signification |
|--------|---------|---------------|
| Confirmée | 🟢 Vert | Réservation validée |
| En attente | 🟠 Orange | En cours de traitement |
| Annulée | 🔴 Rouge | Réservation annulée |
| Soldé | 🟢 Vert | Paiement complet |
| Partiel | 🟠 Orange | Paiement partiel |
| Impayé | 🔴 Rouge | Aucun paiement |
