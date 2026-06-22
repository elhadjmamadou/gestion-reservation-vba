# RESERV+ — Guide d'importation dans Microsoft Access

## Présentation

**RESERV+** est une plateforme de réservation multi-prestataires (Hôtels, Taxis, Vols)  
développée entièrement en Microsoft Access + VBA.

---

## Prérequis

- Microsoft Access 2016 ou version ultérieure (2019, 2021, Microsoft 365)
- Activer les macros et le contenu VBA lors de l'ouverture

---

## Étapes d'installation (à suivre dans l'ordre)

### ÉTAPE 1 — Créer la base de données

1. Ouvrir **Microsoft Access**
2. Cliquer sur **Nouvelle base de données vide**
3. Nommer le fichier : `RESERV_PLUS.accdb`
4. Choisir un emplacement facile d'accès (ex : `C:\RESERV_PLUS\`)
5. Cliquer sur **Créer**

---

### ÉTAPE 2 — Activer le contenu et les macros

1. Si Access affiche une barre jaune **"Contenu désactivé"**, cliquer sur **Activer le contenu**
2. Aller dans **Fichier → Options → Centre de gestion de la confidentialité**
3. Cliquer sur **Paramètres du Centre de gestion de la confidentialité**
4. Sélectionner **Paramètres des macros → Activer toutes les macros**
5. Cocher **Approuver l'accès au modèle objet du projet VBA**
6. Cliquer sur **OK** × 2

---

### ÉTAPE 3 — Ouvrir l'éditeur VBA

Appuyer sur **ALT + F11** pour ouvrir l'éditeur Visual Basic.

---

### ÉTAPE 4 — Importer tous les modules VBA

Dans l'éditeur VBA :

1. Aller dans le menu **Fichier → Importer un fichier...** (ou **Ctrl+M**)
2. Importer les fichiers dans **CET ORDRE EXACT** :

#### 4.1 — Modules Core (importez en premier)
```
vba/core/modConfig.bas
vba/core/modDatabase.bas
vba/core/modSession.bas
vba/core/modSecurite.bas
vba/core/modJournal.bas
vba/core/modUtils.bas
```

#### 4.2 — Modules Métier
```
vba/metier/modClients.bas
vba/metier/modPrestataires.bas
vba/metier/modRessources.bas
vba/metier/modDisponibilite.bas
vba/metier/modReservations.bas
vba/metier/modPaiements.bas
vba/metier/modBillets.bas
vba/metier/modRapports.bas
```

#### 4.3 — Modules UI (interface)
```
vba/ui/modThemeAccess.bas
vba/ui/modNavigation.bas
vba/ui/modFormFactory.bas
vba/ui/modFormsConnexion.bas
vba/ui/modFormsDashboard.bas
vba/ui/modFormsClients.bas
vba/ui/modFormsPrestataires.bas
vba/ui/modFormsRessources.bas
vba/ui/modFormsReservations.bas
vba/ui/modFormsPaiements.bas
vba/ui/modFormsAdministration.bas
vba/ui/modFormsRapports.bas
```

#### 4.4 — Modules Rapports
```
vba/reports/modReportsFactory.bas
vba/reports/modEtatsImprimables.bas
```

#### 4.5 — Modules Installation (importez en dernier)
```
vba/install/modCreationRequetes.bas
vba/install/modInstall_RESERV_PLUS.bas
```

> **ASTUCE** : Pour importer rapidement plusieurs fichiers,  
> dans Windows Explorer, sélectionnez tous les .bas d'un dossier,  
> puis faites glisser-déposer dans l'éditeur VBA.

---

### ÉTAPE 5 — Lancer l'installation

Dans l'éditeur VBA :

1. Appuyer sur **CTRL+G** pour ouvrir la fenêtre Exécution
2. Taper exactement :
   ```
   Installer_RESERV_PLUS
   ```
3. Appuyer sur **ENTRÉE**
4. Cliquer **OK** sur le message de confirmation
5. Patienter ~30 à 60 secondes
6. Un message de succès apparaît à la fin

> Si une erreur apparaît, notez-la et relancez.  
> L'installation gère les objets déjà existants.

---

### ÉTAPE 6 — Vérifier l'installation

Dans la fenêtre Exécution, tapez :
```
TestInstallation
```
Cette procédure vérifie que toutes les tables, formulaires et requêtes ont été créés.

---

### ÉTAPE 7 — Ouvrir l'application

1. Revenir dans Access (ALT+F4 dans VBA ou ALT+TAB)
2. Dans le volet de navigation, double-cliquer sur **F_CONNEXION**
3. Ou dans la fenêtre Exécution : `DoCmd.OpenForm "F_CONNEXION"`

---

### ÉTAPE 8 — Se connecter

| Rôle | Login | Mot de passe |
|------|-------|--------------|
| Administrateur | `admin` | `admin123` |
| Agent | `agent` | `agent123` |
| Utilisateur | `utilisateur` | `user123` |

---

## Résolution des problèmes fréquents

### Erreur "Référence manquante"
- Dans VBA : **Outils → Références**
- Vérifier que **Microsoft DAO 3.6 Object Library** est coché
- Si absent, chercher `DAO360.dll` dans la liste

### Erreur lors de CreateForm / CreateControl
- Ces fonctions nécessitent que la base soit en **mode conception**
- Assurez-vous qu'aucun formulaire n'est ouvert avant l'installation

### Les formulaires sont vides après création
- Fermer et rouvrir Access
- Les formulaires créés par code nécessitent parfois un refresh du volet de navigation

### Erreur "Table déjà existante"
- C'est normal si vous relancez l'installation
- L'installeur vérifie l'existence avant de créer

### Réinstaller depuis zéro
Dans la fenêtre Exécution :
```
SupprimerObjets_RESERV_PLUS
Installer_RESERV_PLUS
```

---

## Structure des fichiers importés

```
RESERV_PLUS_ACCESS/
├── sql/                    ← Référence SQL (non à exécuter manuellement)
├── vba/
│   ├── core/               ← Base : config, session, sécurité, journal
│   ├── metier/             ← Logique : réservations, paiements, billets
│   ├── ui/                 ← Formulaires Access générés par code
│   ├── reports/            ← États imprimables
│   └── install/            ← Installeur principal ← COMMENCER ICI
└── docs/                   ← Documentation complémentaire
```

---

## Contacts et support

Application développée avec Microsoft Access + VBA.  
Consultez `docs/PROCEDURE_TEST_ACCESS.md` pour les tests.
