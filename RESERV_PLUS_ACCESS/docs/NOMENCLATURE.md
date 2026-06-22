# RESERV+ — Nomenclature et conventions

## Tables Access

| Table | Description |
|-------|-------------|
| `PRESTATAIRE` | Propriétaires des ressources |
| `CATEGORIE` | Types : Hôtel, Taxi, Vol |
| `RESSOURCE` | Éléments réservables (table centrale) |
| `CHAMBRE_HOTEL` | Spécialisation hôtel (héritage) |
| `TAXI` | Spécialisation taxi (héritage) |
| `VOL` | Spécialisation vol (héritage) |
| `CLIENT` | Personnes qui réservent |
| `UTILISATEUR` | Comptes de connexion (admin/agent) |
| `RESERVATION` | Réservation centrale |
| `DETAIL_RESERVATION` | Lignes de réservation (table de liaison) |
| `PAIEMENT` | Règlements |
| `BILLET` | Justificatifs émis |
| `JOURNAL_ACTION` | Audit des actions |
| `PARAMETRE` | Configuration application |

## Formulaires Access

| Nom formulaire | Description |
|----------------|-------------|
| `F_CONNEXION` | Écran de connexion |
| `F_TABLEAU_BORD` | Tableau de bord principal |
| `F_CLIENTS` | Gestion des clients |
| `F_PRESTATAIRES` | Gestion des prestataires |
| `F_RESSOURCES` | Catalogue des ressources |
| `F_RESERVATIONS` | Liste et gestion des réservations |
| `F_NOUVELLE_RESERVATION` | Créer une réservation |
| `F_DETAIL_RESERVATION` | Détail complet d'une réservation |
| `F_PAIEMENTS` | Gestion des paiements |
| `F_BILLETS` | Gestion des billets |
| `F_UTILISATEURS` | Administration des comptes |
| `F_PARAMETRES` | Paramètres de l'application |
| `F_ETATS_RAPPORTS` | Rapports et états |

## Requêtes enregistrées

| Nom requête | Description |
|-------------|-------------|
| `RQ_RESERVATIONS_COMPLETES` | Toutes réservations avec client et agent |
| `RQ_RESERVATIONS_DU_JOUR` | Réservations créées aujourd'hui |
| `RQ_RESERVATIONS_EN_ATTENTE` | Réservations statut "En attente" |
| `RQ_TOTAL_ENCAISSE` | Somme de tous les paiements |
| `RQ_SOLDE_PAR_RESERVATION` | Solde restant par réservation |
| `RQ_CLIENTS_AVEC_NB_RESERVATIONS` | Clients avec compteur réservations |
| `RQ_RESSOURCES_AVEC_PRESTATAIRE_CATEGORIE` | Ressources enrichies |
| `RQ_PAIEMENTS_COMPLETS` | Paiements avec détails client |
| `RQ_CHIFFRE_AFFAIRES_PAR_CATEGORIE` | CA groupé par catégorie |
| `RQ_CHIFFRE_AFFAIRES_PAR_PERIODE` | CA par mois/année |
| `RQ_JOURNAL_ACTIONS` | Journal avec noms utilisateurs |
| `RQ_BILLETS_COMPLETS` | Billets avec détails réservation |

## Modules VBA

### core/
| Module | Rôle |
|--------|------|
| `modConfig` | Constantes globales (APP_NOM, rôles, statuts, noms formulaires) |
| `modDatabase` | Utilitaires DAO (ExecuterSQL, ObtenirValeur, TableExiste...) |
| `modSession` | Variables de session globales (g_Login, g_Role...) |
| `modSecurite` | Authentification, contrôle des droits |
| `modJournal` | Journalisation des actions |
| `modUtils` | FormatMontant, FormatDate, validations, utilitaires |

### metier/
| Module | Rôle |
|--------|------|
| `modClients` | CRUD clients |
| `modPrestataires` | CRUD prestataires |
| `modRessources` | CRUD ressources + spécialisations |
| `modDisponibilite` | Vérification des conflits de réservation |
| `modReservations` | Créer, confirmer, annuler réservations |
| `modPaiements` | Enregistrer paiements, calculer soldes |
| `modBillets` | Générer numéros de billet, aperçu |
| `modRapports` | Données pour statistiques et rapports |

### ui/
| Module | Rôle |
|--------|------|
| `modThemeAccess` | Couleurs, polices, fonctions CreateControl |
| `modNavigation` | Fonctions appelées par boutons de navigation |
| `modFormFactory` | Orchestre la création de tous les formulaires |
| `modForms*` | Création et événements de chaque formulaire |

### install/
| Module | Rôle |
|--------|------|
| `modInstall_RESERV_PLUS` | Installation complète (tables, données, formulaires) |
| `modCreationRequetes` | Création des requêtes enregistrées |

## Conventions de nommage

### Variables VBA
- `g_` = variable globale (`g_Login`, `g_Role`)
- `frm` = objet Form
- `rs` = Recordset
- `db` = Database
- `sql` = chaîne SQL
- `id` = identifiant numérique
- `btn` = CommandButton
- `txt` = TextBox
- `lbl` = Label
- `cbo` = ComboBox
- `lst` = ListBox
- `chk` = CheckBox

### Procédures
- `Creer_F_NOM` = création d'un formulaire
- `Creer_RPT_NOM` = création d'un état
- `NomFormulaire_Action` = événement lié à un formulaire (ex: `Clients_Enregistrer`)
- `Navigation_NomModule` = navigation vers un module (appelée par bouton sidebar)

### Couleurs (constantes dans modThemeAccess)
| Constante | Valeur hex | Usage |
|-----------|-----------|-------|
| `COULEUR_SIDEBAR` | #1E2D5A | Fond sidebar navigation |
| `COULEUR_FOND` | #F5F6FA | Fond général des formulaires |
| `COULEUR_PRIMAIRE` | #5B6FBC | Boutons principaux, bandeau |
| `COULEUR_SUCCESS` | #27AE60 | Confirmée, Soldé, succès |
| `COULEUR_WARNING` | #E67E22 | En attente, Partiel, avertissement |
| `COULEUR_DANGER` | #E74C3C | Annulée, Impayé, erreur |
| `COULEUR_TEXTE` | #2C3E50 | Texte principal |
| `COULEUR_BLANC` | #FFFFFF | Fond cartes, texte sur fond sombre |
