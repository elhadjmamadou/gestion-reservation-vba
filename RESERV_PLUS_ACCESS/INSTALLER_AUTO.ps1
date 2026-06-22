# =============================================================================
# RESERV+ — Installeur automatique PowerShell
# Double-cliquez sur INSTALLER_AUTO.bat pour lancer ce script
# =============================================================================

param(
    [string]$CheminDestination = "C:\RESERV_PLUS",
    [string]$NomBase = "RESERV_PLUS.accdb"
)

$ErrorActionPreference = "Stop"

# --- Bannière ---
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║          RESERV+  —  Installation automatique        ║" -ForegroundColor Cyan
Write-Host "  ║     Plateforme de réservation multi-prestataires     ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# --- Chemins ---
$DossierScript = $PSScriptRoot
$DossierVBA    = Join-Path $DossierScript "vba"
$CheminAccdb   = Join-Path $CheminDestination $NomBase

Write-Host "  Source  : $DossierVBA" -ForegroundColor Gray
Write-Host "  Cible   : $CheminAccdb" -ForegroundColor Gray
Write-Host ""

# --- Vérification dossier source ---
if (-not (Test-Path $DossierVBA)) {
    Write-Host "  [ERREUR] Dossier vba\ introuvable : $DossierVBA" -ForegroundColor Red
    Write-Host "  Assurez-vous de lancer ce script depuis le dossier RESERV_PLUS_ACCESS\" -ForegroundColor Yellow
    Read-Host "  Appuyez sur ENTREE pour quitter"
    exit 1
}

# --- Créer le dossier destination ---
if (-not (Test-Path $CheminDestination)) {
    New-Item -ItemType Directory -Path $CheminDestination | Out-Null
    Write-Host "  [OK] Dossier créé : $CheminDestination" -ForegroundColor Green
}

# --- Ordre d'import des modules VBA ---
$Modules = @(
    # CORE (en premier — les autres en dépendent)
    "core\modConfig.bas",
    "core\modDatabase.bas",
    "core\modSession.bas",
    "core\modSecurite.bas",
    "core\modJournal.bas",
    "core\modUtils.bas",

    # METIER
    "metier\modClients.bas",
    "metier\modPrestataires.bas",
    "metier\modRessources.bas",
    "metier\modDisponibilite.bas",
    "metier\modReservations.bas",
    "metier\modPaiements.bas",
    "metier\modBillets.bas",
    "metier\modRapports.bas",

    # UI (thème et navigation avant les formulaires)
    "ui\modThemeAccess.bas",
    "ui\modNavigation.bas",
    "ui\modFormFactory.bas",
    "ui\modFormsConnexion.bas",
    "ui\modFormsDashboard.bas",
    "ui\modFormsClients.bas",
    "ui\modFormsPrestataires.bas",
    "ui\modFormsRessources.bas",
    "ui\modFormsReservations.bas",
    "ui\modFormsPaiements.bas",
    "ui\modFormsAdministration.bas",
    "ui\modFormsRapports.bas",

    # REPORTS
    "reports\modReportsFactory.bas",
    "reports\modEtatsImprimables.bas",

    # INSTALL (en dernier — utilisent tout ce qui précède)
    "install\modCreationRequetes.bas",
    "install\modInstall_RESERV_PLUS.bas"
)

# --- Vérifier que tous les fichiers existent ---
Write-Host "  Vérification des 30 modules VBA..." -ForegroundColor Yellow
$Manquants = @()
foreach ($m in $Modules) {
    $chemin = Join-Path $DossierVBA $m
    if (-not (Test-Path $chemin)) { $Manquants += $m }
}
if ($Manquants.Count -gt 0) {
    Write-Host "  [ERREUR] Fichiers manquants :" -ForegroundColor Red
    $Manquants | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
    Read-Host "  Appuyez sur ENTREE pour quitter"
    exit 1
}
Write-Host "  [OK] 30 fichiers trouvés" -ForegroundColor Green
Write-Host ""

# --- Vérification Trust Access au modèle objet VBA ---
Write-Host "  PRÉREQUIS IMPORTANT" -ForegroundColor Yellow
Write-Host "  ──────────────────────────────────────────────────────" -ForegroundColor Gray
Write-Host "  Dans Access, activez : Fichier → Options → Centre de gestion"  -ForegroundColor White
Write-Host "  de la confidentialité → Paramètres → Paramètres des macros" -ForegroundColor White
Write-Host "  → cocher [v] Approuver l'accès au modèle objet du projet VBA" -ForegroundColor White
Write-Host "  ──────────────────────────────────────────────────────" -ForegroundColor Gray
Write-Host ""

$Reponse = Read-Host "  Ce paramètre est-il activé ? (O pour continuer, N pour annuler)"
if ($Reponse.ToUpper() -ne "O") {
    Write-Host "  Activez d'abord ce paramètre dans Access, puis relancez." -ForegroundColor Yellow
    exit 0
}

# --- Supprimer base existante si demandé ---
if (Test-Path $CheminAccdb) {
    Write-Host ""
    $Rep = Read-Host "  La base $NomBase existe déjà. L'écraser ? (O/N)"
    if ($Rep.ToUpper() -eq "O") {
        Remove-Item $CheminAccdb -Force
        Write-Host "  [OK] Ancienne base supprimée" -ForegroundColor Green
    } else {
        Write-Host "  Installation annulée." -ForegroundColor Yellow
        exit 0
    }
}

# --- Lancer Access via COM ---
Write-Host ""
Write-Host "  Démarrage de Microsoft Access..." -ForegroundColor Cyan
try {
    $Access = New-Object -ComObject Access.Application
} catch {
    Write-Host "  [ERREUR] Microsoft Access n'est pas installé ou est inaccessible." -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Read-Host "  Appuyez sur ENTREE pour quitter"
    exit 1
}

$Access.Visible = $true  # Rendre visible pour suivi visuel

# --- Créer la base de données ---
Write-Host "  Création de $NomBase..." -ForegroundColor Cyan
$Access.NewCurrentDatabase($CheminAccdb)
Start-Sleep -Seconds 2

# --- Importer les modules VBA ---
Write-Host ""
Write-Host "  Import des modules VBA..." -ForegroundColor Cyan
$i = 0
foreach ($m in $Modules) {
    $i++
    $CheminModule = Join-Path $DossierVBA $m
    $NomModule = [System.IO.Path]::GetFileNameWithoutExtension($m)

    try {
        $Access.VBE.ActiveVBProject.VBComponents.Import($CheminModule) | Out-Null
        Write-Host ("  [{0:D2}/{1}] {2}" -f $i, $Modules.Count, $NomModule) -ForegroundColor Green
    } catch {
        Write-Host ("  [{0:D2}/{1}] ERREUR : {2}" -f $i, $Modules.Count, $NomModule) -ForegroundColor Red
        Write-Host "    $_" -ForegroundColor DarkRed
    }

    # Petite pause toutes les 10 imports pour laisser Access respirer
    if ($i % 10 -eq 0) { Start-Sleep -Milliseconds 500 }
}

Write-Host ""
Write-Host "  [$i modules importés]" -ForegroundColor Green

# --- Sauvegarder avant de lancer l'installeur ---
Write-Host ""
Write-Host "  Sauvegarde..." -ForegroundColor Cyan
$Access.CurrentDb().Close()
Start-Sleep -Seconds 1
$Access.OpenCurrentDatabase($CheminAccdb)
Start-Sleep -Seconds 2

# --- Lancer l'installeur RESERV+ ---
Write-Host "  Lancement de Installer_RESERV_PLUS..." -ForegroundColor Cyan
Write-Host "  (création des 14 tables, 12 requêtes, 11 formulaires, données démo)" -ForegroundColor Gray

try {
    $Access.Run("Installer_RESERV_PLUS")
    Write-Host ""
    Write-Host "  [OK] Installation terminée avec succès !" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "  [ATTENTION] Erreur lors de l'installation automatique :" -ForegroundColor Yellow
    Write-Host "  $_" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  → Ouvrez la fenêtre Exécution VBA (CTRL+G) et tapez manuellement :" -ForegroundColor Cyan
    Write-Host "     Installer_RESERV_PLUS" -ForegroundColor White
}

# --- Résultat final ---
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║               INSTALLATION TERMINÉE                 ║" -ForegroundColor Green
Write-Host "  ╠══════════════════════════════════════════════════════╣" -ForegroundColor Green
Write-Host "  ║  Connexion admin  : admin     / admin123             ║" -ForegroundColor Green
Write-Host "  ║  Connexion agent  : agent     / agent123             ║" -ForegroundColor Green
Write-Host "  ║  Connexion user   : utilisateur / user123            ║" -ForegroundColor Green
Write-Host "  ╠══════════════════════════════════════════════════════╣" -ForegroundColor Green
Write-Host "  ║  Base créée dans  : $CheminDestination$('' * [Math]::Max(0, 22 - $CheminDestination.Length)) ║" -ForegroundColor Green
Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""

Read-Host "  Appuyez sur ENTREE pour terminer"
