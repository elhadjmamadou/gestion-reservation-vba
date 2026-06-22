# RESERV+ — Règles métier

## 1. Connexion et sessions

- L'identifiant et le mot de passe sont vérifiés dans la table `UTILISATEUR`
- Les comptes inactifs (`actif = False`) sont refusés
- La session est stockée dans des variables globales VBA : `g_IdUtilisateur`, `g_Login`, `g_Role`, `g_NomComplet`
- Chaque connexion est enregistrée dans `JOURNAL_ACTION`
- La session s'efface à la déconnexion ou à la fermeture d'Access

## 2. Rôles et droits

| Action | Administrateur | Agent | Utilisateur |
|--------|:--------------:|:-----:|:-----------:|
| Voir tableau de bord | ✓ | ✓ | ✓ |
| Voir réservations | ✓ | ✓ | ✓ |
| Créer réservation | ✓ | ✓ | ✗ |
| Annuler réservation | ✓ | ✓ | ✗ |
| Enregistrer paiement | ✓ | ✓ | ✗ |
| Générer billet | ✓ | ✓ | ✗ |
| Gérer clients | ✓ | ✓ | Vue seule |
| Gérer prestataires | ✓ | ✓ | Vue seule |
| Gérer ressources | ✓ | ✓ | Vue seule |
| Supprimer client/ressource | ✓ | ✗ | ✗ |
| Gérer utilisateurs | ✓ | ✗ | ✗ |
| Modifier paramètres | ✓ | ✗ | ✗ |
| Voir journal des actions | ✓ | ✗ | ✗ |
| Vider journal | ✓ | ✗ | ✗ |

## 3. Disponibilité des ressources

**Règle de chevauchement** : Une ressource est en conflit si une réservation **non annulée** existe avec :
```
date_debut_existante < nouvelle_date_fin
ET
date_fin_existante > nouvelle_date_debut
```

- Le champ `disponible` (YESNO) indique une indisponibilité permanente (maintenance, fermeture)
- Une ressource marquée `disponible = False` est toujours indisponible, même sans réservation
- Les réservations annulées (`statut = 'Annulée'`) ne bloquent pas la disponibilité

## 4. Calcul des montants

```
durée = DateDiff("d", date_debut, date_fin)
si durée <= 0 → durée = 1

sous_total = durée × prix_unitaire_appliqué × quantité
montant_total = ΣSous-totaux de DETAIL_RESERVATION
```

- Le prix est **figé au moment de la réservation** dans `prix_unitaire_applique`
- Une modification du prix de la ressource n'affecte pas les réservations existantes
- Le `montant_total` est recalculé après chaque ajout/suppression de détail

## 5. Gestion des paiements

```
total_payé = ΣMontants de PAIEMENT pour cette réservation
solde = montant_total - total_payé

État :
  total_payé = 0           → "Impayé"
  0 < total_payé < montant → "Partiel"
  total_payé >= montant    → "Soldé"
```

- Les paiements partiels sont autorisés
- Un paiement peut dépasser le montant total (avec avertissement)
- Impossible d'enregistrer un paiement pour une réservation annulée
- Seul un administrateur peut supprimer un paiement

## 6. Numérotation des billets

**Format** : `PREFIXE-YYYYMMDD-000001`

- Le préfixe est lu depuis `PARAMETRE` (clé = `PREFIXE_BILLET`, défaut = `BIL`)
- La séquence repart de 000001 chaque jour
- L'unicité est garantie par une vérification avant insertion
- Si un billet actif existe pour une réservation, il est annulé lors de la génération d'un nouveau

## 7. Héritage des ressources

Pattern **une table par type** :
- `RESSOURCE` = attributs communs (libellé, prix, capacité, prestataire, catégorie)
- `CHAMBRE_HOTEL` (clé = `id_ressource`) = étoiles, type chambre, lits, petit-déjeuner, vue
- `TAXI` (clé = `id_ressource`) = marque, modèle, immatriculation, places, climatisation
- `VOL` (clé = `id_ressource`) = numéro vol, villes, dates/heures, classe

L'`id_ressource` est partagé : supprimer une ressource supprime aussi son enregistrement spécialisé.

## 8. Journal des actions

Toute action importante est enregistrée :
- Connexion / Déconnexion
- Création / Modification / Suppression de données
- Génération de billets et reçus
- Consultation des rapports
- Modifications des paramètres

Le journal ne bloque jamais l'application si une erreur d'écriture survient.

## 9. Contraintes de suppression

| Entité | Condition de suppression |
|--------|--------------------------|
| Client | Aucune réservation liée |
| Prestataire | Aucune ressource rattachée |
| Ressource | Aucune réservation active (non annulée) |
| Utilisateur | Ne pas être l'utilisateur connecté |
| Réservation | Admin seulement, après confirmation double |
