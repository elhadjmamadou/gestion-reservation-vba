# RESERV+ — Changelog

## v1.0.0 — Juin 2026 (Version initiale)

### Ajouté
- 14 tables Access : PRESTATAIRE, CATEGORIE, RESSOURCE, CHAMBRE_HOTEL, TAXI, VOL,
  CLIENT, UTILISATEUR, RESERVATION, DETAIL_RESERVATION, PAIEMENT, BILLET,
  JOURNAL_ACTION, PARAMETRE
- 12 requêtes enregistrées (RQ_*)
- 11 formulaires créés par code VBA (CreateForm + CreateControl)
- Thème visuel cohérent : sidebar marine, cartes statistiques, statuts colorés
- Authentification : login + mot de passe + vérification compte actif
- 3 niveaux de rôles : Administrateur, Agent, Utilisateur
- Vérification de disponibilité par détection de chevauchement de dates
- Calcul automatique : durée × prix × quantité
- Gestion des paiements partiels avec calcul de solde
- Génération de billets numérotés BIL-YYYYMMDD-000001
- Journal des actions complet
- Module d'installation automatique (Installer_RESERV_PLUS)
- Procédure de test (TestInstallation)
- Données de démonstration : 3 prestataires, 3 ressources, 5 clients
- Export Excel via DoCmd.TransferSpreadsheet
- Rapports : chiffre d'affaires, taux d'occupation, journal

### Architecture
- Pattern héritage "une table par type" pour RESSOURCE
- Table de liaison DETAIL_RESERVATION pour l'association n,n
- Variables de session VBA globales (modSession)
- Modules séparés : core / metier / ui / install / reports
