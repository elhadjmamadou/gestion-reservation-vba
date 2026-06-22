# RESERV+ — Résumé du Modèle Conceptuel de Données

## Entités et rôles

| Entité | Rôle |
|--------|------|
| **PRESTATAIRE** | Propriétaire qui confie ses ressources à la plateforme |
| **CATEGORIE** | Type général de ressource (Hôtel, Taxi, Vol) |
| **RESSOURCE** | Élément réservable rattaché à un prestataire |
| **CHAMBRE_HOTEL** | Spécialisation hôtelière (héritage de RESSOURCE) |
| **TAXI** | Spécialisation taxi (héritage de RESSOURCE) |
| **VOL** | Spécialisation vol (héritage de RESSOURCE) |
| **CLIENT** | Personne qui effectue les réservations |
| **UTILISATEUR** | Agent/Admin qui gère l'application (login) |
| **RESERVATION** | Engagement central reliant client et ressource(s) |
| **DETAIL_RESERVATION** | Table de liaison RESERVATION ↔ RESSOURCE (n,n) |
| **PAIEMENT** | Règlement rattaché à une réservation |
| **BILLET** | Justificatif émis pour une réservation |
| **JOURNAL_ACTION** | Traçabilité des opérations |
| **PARAMETRE** | Configuration de l'application |

## Associations et cardinalités

| Association | Entités | Cardinalités | Implémentation |
|-------------|---------|--------------|----------------|
| POSSEDER | PRESTATAIRE — RESSOURCE | 1,n / 1,1 | FK `id_prestataire` dans RESSOURCE |
| CLASSER | CATEGORIE — RESSOURCE | 1,n / 1,1 | FK `id_categorie` dans RESSOURCE |
| SPECIALISER | RESSOURCE — CHAMBRE/TAXI/VOL | 1,1 / 0,1 | PK partagée (`id_ressource`) |
| EFFECTUER | CLIENT — RESERVATION | 1,n / 1,1 | FK `id_client` dans RESERVATION |
| SAISIR | UTILISATEUR — RESERVATION | 1,n / 1,1 | FK `id_utilisateur` dans RESERVATION |
| CONCERNER | RESERVATION — RESSOURCE | n,n | Table DETAIL_RESERVATION |
| REGLER | RESERVATION — PAIEMENT | 1,n / 1,1 | FK `id_reservation` dans PAIEMENT |
| GENERER | RESERVATION — BILLET | 1,1 / 1,1 | FK `id_reservation` dans BILLET |

## Choix de conception

### Héritage "une table par type"
Les spécialisations (CHAMBRE_HOTEL, TAXI, VOL) partagent la PK `id_ressource` avec RESSOURCE.  
Avantages : pas de colonnes vides, évolution indépendante, requêtes propres.

### Prestataires sans compte
Les prestataires sont des fiches de données. Seuls les agents/admins se connectent.  
Simplification : pas de gestion de droits pour les prestataires.

### Association n,n via DETAIL_RESERVATION
Une réservation peut porter sur plusieurs ressources (ex: chambre + taxi).  
La table DETAIL_RESERVATION stocke : quantité, prix appliqué (figé), sous-total.

## Schéma simplifié

```
PRESTATAIRE ──── RESSOURCE ──── CATEGORIE
                     │
         ┌───────────┼───────────┐
    CHAMBRE_HOTEL   TAXI        VOL

CLIENT ──── RESERVATION ──── UTILISATEUR
                │
        DETAIL_RESERVATION ─── RESSOURCE
                │
          ┌─────┴─────┐
       PAIEMENT     BILLET
```
