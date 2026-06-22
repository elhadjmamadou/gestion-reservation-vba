# RESERV+ — Démarrage rapide (3 méthodes)

## Choisissez votre méthode

| Méthode | Temps | Prérequis |
|---------|-------|-----------|
| **A — Script PowerShell** (recommandée) | ~2 min | Windows 10/11, Access ouvert |
| **B — Bootstrap VBA** (1 seul import) | ~3 min | Activer "Approuver VBA" dans Access |
| **C — Import manuel** (classique) | ~10 min | Aucun |

---

## MÉTHODE A — Script PowerShell (la plus rapide)

### Prérequis une seule fois dans Access
1. **Fichier → Options → Centre de gestion de la confidentialité**
2. **Paramètres → Paramètres des macros**
3. Cocher **[v] Approuver l'accès au modèle objet du projet VBA**
4. Cliquer **OK × 2**

### Installation
1. Copier le dossier `RESERV_PLUS_ACCESS\` sur votre PC Windows (ex: `C:\RESERV_PLUS_ACCESS\`)
2. **Double-cliquer sur `INSTALLER_AUTO.bat`**
3. Répondre **O** quand demandé
4. Access s'ouvre, 30 modules s'importent, l'installeur se lance automatiquement
5. Se connecter : `admin` / `admin123`

---

## MÉTHODE B — Bootstrap VBA (si PowerShell est bloqué)

### Étape 1 — Un seul import manuel
1. Ouvrir Access → créer `RESERV_PLUS.accdb`
2. **ALT+F11** → ouvrir l'éditeur VBA
3. **CTRL+M** → importer **uniquement** : `vba\BOOTSTRAP_IMPORT.bas`

### Étape 2 — Tout importer d'un coup
1. **CTRL+G** → fenêtre Exécution
2. Taper :
   ```
   LancerInstallationComplete
   ```
3. Entrer le chemin du dossier `vba\` quand demandé  
   Exemple : `C:\RESERV_PLUS_ACCESS\vba`
4. Cliquer **Oui** pour lancer l'installation
5. Se connecter : `admin` / `admin123`

---

## MÉTHODE C — Import manuel classique

Suivez le fichier `README_IMPORTATION_ACCESS.md` (8 étapes détaillées avec résolution de problèmes).

---

## Après l'installation — Vérification rapide

Dans la fenêtre Exécution VBA (CTRL+G) :
```
TestInstallation
```
Résultat attendu :
```
Tables  : 14/14 ✔
Formes  : 11/11 ✔
Requêtes: OK    ✔
```

---

## Comptes démo

| Rôle | Login | Mot de passe |
|------|-------|--------------|
| Administrateur | `admin` | `admin123` |
| Agent | `agent` | `agent123` |
| Utilisateur | `utilisateur` | `user123` |
