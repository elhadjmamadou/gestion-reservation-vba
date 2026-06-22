/* ============================================================
   RESERV+ - Données initiales
   Fichier : 02_donnees_initiales.sql
   ============================================================ */

/* ---- Catégories de base ---- */
INSERT INTO CATEGORIE (libelle_categorie, description)
VALUES ('Hôtel', 'Chambres et hébergements hôteliers');

INSERT INTO CATEGORIE (libelle_categorie, description)
VALUES ('Taxi', 'Véhicules de transport terrestre');

INSERT INTO CATEGORIE (libelle_categorie, description)
VALUES ('Vol', 'Vols et transport aérien');

/* ---- Utilisateurs initiaux ---- */
INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation)
VALUES ('admin', 'admin123', 'Administrateur Système', 'Administrateur', True, Now());

INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation)
VALUES ('agent', 'agent123', 'Agent de Réservation', 'Agent', True, Now());

INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation)
VALUES ('utilisateur', 'user123', 'Utilisateur Standard', 'Utilisateur', True, Now());

/* ---- Paramètres de l'application ---- */
INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description)
VALUES ('NOM_APPLICATION', 'RESERV+', 'Nom de l application');

INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description)
VALUES ('DEVISE', 'GNF', 'Devise utilisée pour les montants');

INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description)
VALUES ('PREFIXE_BILLET', 'BIL', 'Préfixe du numéro de billet');

INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description)
VALUES ('NOM_ETABLISSEMENT', 'RESERV+', 'Nom affiché sur les documents');

INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description)
VALUES ('VERSION_APP', '1.0.0', 'Version de l application');

INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description)
VALUES ('FREQUENCE_SAUVEGARDE', '7', 'Fréquence de sauvegarde en jours');

/* ---- Prestataires de démonstration ---- */
INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion)
VALUES ('Hôtel Kaloum SARL', 'M. Sylla', 'contact@hotelkaloum.gn', '620 11 22 33',
        'Avenue de la République, Kaloum, Conakry', 'Hôtellerie', Now());

INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion)
VALUES ('Express Taxi Conakry', 'Mme Barry', 'info@expresstaxi.gn', '622 33 44 55',
        'Quartier Madina, Conakry', 'Transport terrestre', Now());

INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion)
VALUES ('Air Guinée Express', 'M. Camara', 'reservations@airgn.gn', '628 55 66 77',
        'Aéroport International Gbessia, Conakry', 'Aérien', Now());

INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion)
VALUES ('Niger Hôtels Group', 'M. Diakité', 'contact@nigerhotels.gn', '664 77 88 99',
        'Avenue du Niger, Conakry', 'Hôtellerie', Now());

/* ---- Ressources de démonstration ---- */
INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation)
VALUES (1, 1, 'Hôtel Kaloum - Chambre 204', 'Chambre double vue mer, climatisée', 150000, 2, True, 'Disponible', Now());

INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation)
VALUES (4, 1, 'Hôtel Niger - Chambre 11', 'Chambre standard confortable', 100000, 2, True, 'Disponible', Now());

INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation)
VALUES (2, 2, 'Taxi Express 01', 'Berline climatisée 4 places', 30000, 4, True, 'Disponible', Now());

INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation)
VALUES (2, 2, 'Taxi VIP 07', 'SUV climatisé 4 places, haut de gamme', 60000, 4, True, 'Disponible', Now());

INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation)
VALUES (3, 3, 'Vol CKY-LBE', 'Conakry - Labé, classe économique', 620000, 120, True, 'Disponible', Now());

INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation)
VALUES (3, 3, 'Vol CKY-NZE', 'Conakry - N''Zérékoré, classe économique', 580000, 120, True, 'Disponible', Now());

/* ---- Détails spécifiques CHAMBRE_HOTEL ---- */
INSERT INTO CHAMBRE_HOTEL (id_ressource, nb_etoiles, type_chambre, nb_lits, petit_dejeuner, vue)
VALUES (1, 4, 'Double', 2, True, 'Vue mer');

INSERT INTO CHAMBRE_HOTEL (id_ressource, nb_etoiles, type_chambre, nb_lits, petit_dejeuner, vue)
VALUES (2, 3, 'Standard', 1, False, 'Vue jardin');

/* ---- Détails spécifiques TAXI ---- */
INSERT INTO TAXI (id_ressource, marque, modele, immatriculation, nb_places, climatisation)
VALUES (3, 'Toyota', 'Corolla', 'RC-1234-CKY', 4, True);

INSERT INTO TAXI (id_ressource, marque, modele, immatriculation, nb_places, climatisation)
VALUES (4, 'Toyota', 'Land Cruiser', 'RC-5678-CKY', 4, True);

/* ---- Détails spécifiques VOL ---- */
INSERT INTO VOL (id_ressource, num_vol, ville_depart, ville_arrivee, date_depart, heure_depart, date_arrivee, heure_arrivee, classe)
VALUES (5, 'AGX-001', 'Conakry', 'Labé', #2026/06/25#, #2026/06/25 08:00:00#, #2026/06/25#, #2026/06/25 10:30:00#, 'Économique');

INSERT INTO VOL (id_ressource, num_vol, ville_depart, ville_arrivee, date_depart, heure_depart, date_arrivee, heure_arrivee, classe)
VALUES (6, 'AGX-002', 'Conakry', 'N''Zérékoré', #2026/06/26#, #2026/06/26 09:00:00#, #2026/06/26#, #2026/06/26 11:45:00#, 'Économique');

/* ---- Clients de démonstration ---- */
INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription)
VALUES ('Diallo', 'Mamadou', 'm.diallo@mail.gn', '620 00 00 01', 'Kaloum, Conakry', #2026/01/02#);

INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription)
VALUES ('Camara', 'Aïssatou', 'a.camara@mail.gn', '622 11 22 33', 'Ratoma, Conakry', #2026/02/15#);

INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription)
VALUES ('Bah', 'Saikou', 's.bah@mail.gn', '628 44 55 66', 'Matoto, Conakry', #2026/03/20#);

INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription)
VALUES ('Touré', 'Fatou', 'f.toure@mail.gn', '655 77 88 99', 'Dixinn, Conakry', #2026/04/10#);

INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription)
VALUES ('Soumah', 'Karim', 'k.soumah@mail.gn', '664 12 34 56', 'Matam, Conakry', #2026/05/05#);
