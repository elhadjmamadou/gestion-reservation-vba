# RESERV+ — Procédure de test dans Access

## Prérequis

Installation complète effectuée. Connectez-vous avec `admin / admin123`.

---

## Test 1 — Vérification de l'installation

Dans la fenêtre VBA (Ctrl+G) :
```
TestInstallation
```
**Résultat attendu :** Message affichant 14/14 tables, 11/11 formulaires, requêtes OK.

---

## Test 2 — Connexion et rôles

### 2a. Connexion Administrateur
1. Ouvrir `F_CONNEXION`
2. Login : `admin` | MDP : `admin123`
3. Cliquer **SE CONNECTER**
4. **Résultat attendu :** Tableau de bord ouvert, tous les menus visibles

### 2b. Connexion Agent
1. Se déconnecter (bouton X dans la sidebar)
2. Login : `agent` | MDP : `agent123`
3. **Résultat attendu :** Tableau de bord ouvert, menu Paramètres masqué

### 2c. Mauvais mot de passe
1. Saisir `admin` + mot de passe erroné
2. **Résultat attendu :** Message d'erreur rouge, champ MDP vidé

---

## Test 3 — Créer un client

1. Aller dans **Clients** (sidebar ou bouton)
2. Cliquer **+ Nouveau client**
3. Saisir :
   - Nom : `Barry`
   - Prénom : `Ibrahima`
   - Téléphone : `622 55 66 77`
   - Email : `i.barry@test.gn`
4. Cliquer **Enregistrer**
5. **Résultat attendu :** Client apparaît dans la liste, compteur incrémenté

---

## Test 4 — Créer un prestataire

1. Aller dans **Prestataires**
2. Cliquer **+ Nouveau prestataire**
3. Saisir :
   - Nom entreprise : `Hôtel Mamou`
   - Type activité : `Hôtellerie`
   - Responsable : `M. Diallo`
   - Téléphone : `655 00 11 22`
4. Cliquer **Enregistrer**
5. **Résultat attendu :** Prestataire affiché dans la liste

---

## Test 5 — Créer une ressource Hôtel

1. Aller dans **Ressources**
2. Cliquer **+ Nouvelle ressource**
3. Saisir :
   - Prestataire : `Hôtel Mamou`
   - Catégorie : `Hôtel`
   - Libellé : `Chambre Standard 101`
   - Prix unitaire : `80000`
   - Capacité : `2`
4. Cliquer **Enregistrer**
5. **Résultat attendu :** Ressource dans la liste, statut Disponible

---

## Test 6 — Créer une réservation

1. Aller dans **Réservations → + Nouvelle réservation**
2. Sélectionner :
   - Client : `Barry Ibrahima`
   - Catégorie : `Hôtel`
   - Ressource : `Chambre Standard 101`
   - Date début : demain
   - Date fin : dans 3 jours
   - Nb personnes : 2
3. Cliquer **Vérifier disponibilité**
4. **Résultat attendu :** Bandeau vert "Disponible — Aucun conflit"
5. Vérifier le montant calculé : `80 000 × 2 jours = 160 000 GNF`
6. Cliquer **Enregistrer**
7. **Résultat attendu :** Message succès + billet généré automatiquement

---

## Test 7 — Vérifier le conflit de disponibilité

1. Créer une 2ème réservation pour la même chambre, même période
2. Cliquer **Vérifier disponibilité**
3. **Résultat attendu :** Bandeau rouge "Indisponible — Conflit détecté"
4. Tenter d'**Enregistrer** quand même
5. **Résultat attendu :** Message d'erreur bloquant la création

---

## Test 8 — Enregistrer un paiement partiel

1. Aller dans **Paiements**
2. Sélectionner la réservation créée au test 6
3. Vérifier le solde affiché (160 000 GNF)
4. Saisir :
   - Montant : `80 000`
   - Mode : `Espèces`
5. Cliquer **$ Encaisser**
6. **Résultat attendu :** Paiement enregistré, liste mise à jour

---

## Test 9 — Vérifier le solde restant

1. Dans **Paiements**, vérifier la colonne Solde
2. **Résultat attendu :** Solde = 80 000 GNF, État = "Partiel"
3. Enregistrer un 2ème paiement de 80 000 GNF
4. **Résultat attendu :** Solde = 0 GNF, État = "Soldé"

---

## Test 10 — Générer un billet

### Via VBA (fenêtre Exécution) :
```
ApercuBillet 1
```
**Résultat attendu :** Fenêtre texte avec le billet formaté.

### Via le formulaire Réservations :
1. Sélectionner une réservation dans la liste
2. Cliquer **Billet**
3. **Résultat attendu :** Message de succès avec le numéro BIL-YYYYMMDD-000001

---

## Test 11 — Tableau de bord

1. Aller dans **Tableau de bord**
2. Cliquer **Actualiser**
3. **Résultat attendu :**
   - Compteur réservations > 0
   - Compteur du jour > 0
   - Montant encaissé > 0
   - Dernières réservations affichées

---

## Test 12 — Rapport chiffre d'affaires

1. Aller dans **États & rapports**
2. Cliquer **$ Chiffre d'affaires**
3. **Résultat attendu :** Liste par catégorie avec totaux

---

## Test 13 — Rapport journalier (VBA)

Dans la fenêtre Exécution :
```
RapportJournalier
```
**Résultat attendu :** Synthèse complète du jour avec totaux.

---

## Test 14 — Journal des actions

1. Aller dans **États & rapports**
2. Cliquer **Journal actions** (admin seulement)
3. **Résultat attendu :** Toutes les actions depuis l'installation

---

## Test 15 — Annuler une réservation

1. Dans **Réservations**, sélectionner une réservation
2. Cliquer **Annuler**
3. Confirmer
4. **Résultat attendu :** Statut passe à "Annulée" (rouge)
5. Créer une nouvelle réservation sur le même créneau
6. **Résultat attendu :** La vérification de disponibilité dit "Disponible"  
   (les réservations annulées ne bloquent plus)

---

## Test 16 — Gestion des utilisateurs (admin)

1. Aller dans **Paramètres → Gérer utilisateurs**
2. Créer un nouvel utilisateur :
   - Login : `test_user`
   - MDP : `test123`
   - Rôle : `Agent`
3. Se déconnecter et se reconnecter avec `test_user / test123`
4. **Résultat attendu :** Connexion réussie avec rôle Agent

---

## Résultats attendus globaux

| Test | Action | Résultat |
|------|--------|----------|
| Installation | TestInstallation | 14 tables, 11 formulaires |
| Connexion admin | admin/admin123 | Dashboard complet |
| Connexion agent | agent/agent123 | Dashboard sans admin |
| Mauvais MDP | login erroné | Message d'erreur |
| Créer client | Formulaire clients | Client dans la liste |
| Créer réservation | Formulaire réservation | Montant calculé auto |
| Dispo libre | Vérifier | Bandeau vert |
| Dispo occupée | Vérifier | Bandeau rouge + blocage |
| Paiement partiel | Encaisser 50% | Solde = 50%, état Partiel |
| Solde nul | Encaisser 100% | État Soldé |
| Billet | Générer | Numéro BIL-YYYYMMDD-000001 |
| Annulation | Annuler réservation | Statut Annulée, créneau libéré |
