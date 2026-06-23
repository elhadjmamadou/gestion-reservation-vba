# RESERV+ - Installeur automatique PowerShell (ASCII only)
# Lancer via INSTALLER_AUTO.bat

param(
    [string]$CheminDestination = "C:\RESERV_PLUS",
    [string]$NomBase = "RESERV_PLUS.accdb"
)

$ErrorActionPreference = "Stop"

Clear-Host
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  RESERV+  -  Installation automatique" -ForegroundColor Cyan
Write-Host "  Plateforme de reservation multi-prestataires" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

$DossierScript = $PSScriptRoot
$DossierVBA    = Join-Path $DossierScript "vba"
$CheminAccdb   = Join-Path $CheminDestination $NomBase

Write-Host "  Source  : $DossierVBA" -ForegroundColor Gray
Write-Host "  Cible   : $CheminAccdb" -ForegroundColor Gray
Write-Host ""

# Verification dossier source
if (-not (Test-Path $DossierVBA)) {
    Write-Host "  [ERREUR] Dossier vba\ introuvable : $DossierVBA" -ForegroundColor Red
    Write-Host "  Lancez ce script depuis le dossier RESERV_PLUS_ACCESS\" -ForegroundColor Yellow
    Read-Host "  Appuyez sur ENTREE pour quitter"
    exit 1
}

# Creer le dossier destination
if (-not (Test-Path $CheminDestination)) {
    New-Item -ItemType Directory -Path $CheminDestination | Out-Null
    Write-Host "  [OK] Dossier cree : $CheminDestination" -ForegroundColor Green
}

# Ordre d'import obligatoire
$Modules = @(
    "core\modConfig.bas",
    "core\modDatabase.bas",
    "core\modSession.bas",
    "core\modSecurite.bas",
    "core\modJournal.bas",
    "core\modUtils.bas",
    "metier\modClients.bas",
    "metier\modPrestataires.bas",
    "metier\modRessources.bas",
    "metier\modDisponibilite.bas",
    "metier\modReservations.bas",
    "metier\modPaiements.bas",
    "metier\modBillets.bas",
    "metier\modRapports.bas",
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
    "reports\modReportsFactory.bas",
    "reports\modEtatsImprimables.bas",
    "install\modCreationRequetes.bas",
    "install\modInstall_RESERV_PLUS.bas"
)

# Verification que tous les fichiers existent
Write-Host "  Verification des 30 modules VBA..." -ForegroundColor Yellow
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
Write-Host "  [OK] 30 fichiers trouves" -ForegroundColor Green
Write-Host ""

# Prerequis VBA
Write-Host "  PREREQUIS IMPORTANT" -ForegroundColor Yellow
Write-Host "  ----------------------------------------------------------" -ForegroundColor Gray
Write-Host "  Dans Access : Fichier > Options > Centre de gestion" -ForegroundColor White
Write-Host "  de la confidentialite > Parametres des macros" -ForegroundColor White
Write-Host "  > cocher [v] Approuver l'acces au modele objet du projet VBA" -ForegroundColor White
Write-Host "  ----------------------------------------------------------" -ForegroundColor Gray
Write-Host ""

$Reponse = Read-Host "  Ce parametre est-il active ? (O pour continuer, N pour annuler)"
if ($Reponse.ToUpper() -ne "O") {
    Write-Host "  Activez ce parametre dans Access, puis relancez." -ForegroundColor Yellow
    exit 0
}

# Supprimer base existante
if (Test-Path $CheminAccdb) {
    Write-Host ""
    $Rep = Read-Host "  La base $NomBase existe deja. L'ecraser ? (O/N)"
    if ($Rep.ToUpper() -eq "O") {
        Remove-Item $CheminAccdb -Force
        Write-Host "  [OK] Ancienne base supprimee" -ForegroundColor Green
    } else {
        Write-Host "  Installation annulee." -ForegroundColor Yellow
        exit 0
    }
}

# Lancer Access via COM
Write-Host ""
Write-Host "  Demarrage de Microsoft Access..." -ForegroundColor Cyan
try {
    $Access = New-Object -ComObject Access.Application
} catch {
    Write-Host "  [ERREUR] Microsoft Access inaccessible." -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Read-Host "  Appuyez sur ENTREE pour quitter"
    exit 1
}

$Access.Visible = $true

# Creer la base de donnees
Write-Host "  Creation de $NomBase..." -ForegroundColor Cyan
$Access.NewCurrentDatabase($CheminAccdb)
Start-Sleep -Seconds 2

# Importer les modules VBA
Write-Host ""
Write-Host "  Import des 30 modules VBA..." -ForegroundColor Cyan
$i = 0
foreach ($m in $Modules) {
    $i++
    $CheminModule = Join-Path $DossierVBA $m
    $NomModule = [System.IO.Path]::GetFileNameWithoutExtension($m)

    try {
        $Access.VBE.ActiveVBProject.VBComponents.Import($CheminModule) | Out-Null
        Write-Host ("  [{0:D2}/{1}] OK : {2}" -f $i, $Modules.Count, $NomModule) -ForegroundColor Green
    } catch {
        Write-Host ("  [{0:D2}/{1}] ERREUR : {2} -> {3}" -f $i, $Modules.Count, $NomModule, $_.Exception.Message) -ForegroundColor Red
    }

    if ($i % 10 -eq 0) { Start-Sleep -Milliseconds 500 }
}

Write-Host ""
Write-Host "  [$i modules importes]" -ForegroundColor Green

# Sauvegarder et rouvrir
Write-Host ""
Write-Host "  Sauvegarde..." -ForegroundColor Cyan
try { $Access.CurrentDb().Close() } catch {}
Start-Sleep -Seconds 1
$Access.OpenCurrentDatabase($CheminAccdb)
Start-Sleep -Seconds 2

# Lancer l'installeur
Write-Host "  Lancement de Installer_RESERV_PLUS..." -ForegroundColor Cyan
Write-Host "  (14 tables, 12 requetes, 11 formulaires, donnees demo)" -ForegroundColor Gray

try {
    $Access.Run("Installer_RESERV_PLUS")
    Write-Host ""
    Write-Host "  [OK] Installation terminee avec succes !" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "  [INFO] Lancez manuellement dans Access (CTRL+G) :" -ForegroundColor Yellow
    Write-Host "         Installer_RESERV_PLUS" -ForegroundColor White
    Write-Host "  Erreur : $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# Resultat final
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  INSTALLATION TERMINEE" -ForegroundColor Green
Write-Host "  Connexion admin      : admin       / admin123" -ForegroundColor Green
Write-Host "  Connexion agent      : agent       / agent123" -ForegroundColor Green
Write-Host "  Connexion utilisateur: utilisateur / user123" -ForegroundColor Green
Write-Host "  Base creee dans      : $CheminDestination" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""

Read-Host "  Appuyez sur ENTREE pour terminer"
