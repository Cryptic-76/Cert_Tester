# =====================================================================
# Interaktives Zertifikats-Tool
# Backup, Restore, Prüfung, HTML-Report, Cleanup
# =====================================================================

$global:BackupRoot = "C:\ZertifikatBackup"
$timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")

# -------------------------------
# Zertifikat-Stores
# -------------------------------
$global:Stores = @(
    "Cert:\LocalMachine\My",
    "Cert:\LocalMachine\Root",
    "Cert:\LocalMachine\CA",
    "Cert:\LocalMachine\TrustedPublisher",
    "Cert:\LocalMachine\TrustedPeople",
    "Cert:\CurrentUser\My",
    "Cert:\CurrentUser\Root",
    "Cert:\CurrentUser\CA",
    "Cert:\CurrentUser\TrustedPublisher",
    "Cert:\CurrentUser\TrustedPeople"
)

# =====================================================================
# FUNKTION 1: BACKUP ERSTELLEN
# =====================================================================
function Create-Backup {
    $backupPath = "$BackupRoot\Backup_$timestamp"
    New-Item -ItemType Directory -Path $backupPath -Force | Out-Null

    foreach ($store in $Stores) {
        $storeName = $store.Replace("Cert:\LocalMachine\", "LM_").Replace("Cert:\CurrentUser\", "CU_")
        $exportDir = "$backupPath\$storeName"
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null

        Get-ChildItem $store | ForEach-Object {
            $fileName = "$($_.Thumbprint).cer"
            Export-Certificate -Cert $_ -FilePath "$exportDir\$fileName" -Force | Out-Null
        }
    }

    Write-Host "Backup erstellt unter: $backupPath" -ForegroundColor Green
}

# =====================================================================
# FUNKTION 2: BACKUP ZURÜCKSPIELEN
# =====================================================================
function Restore-Backup {
    $path = Read-Host "Bitte Backup-Pfad eingeben"

    if (-not (Test-Path $path)) {
        Write-Host "Pfad existiert nicht." -ForegroundColor Red
        return
    }

    Get-ChildItem -Recurse -Path $path -Filter *.cer | ForEach-Object {
        Write-Host "Importiere: $($_.FullName)"
        Import-Certificate -FilePath $_.FullName -CertStoreLocation Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Out-Null
        Import-Certificate -FilePath $_.FullName -CertStoreLocation Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Out-Null
    }

    Write-Host "Backup erfolgreich zurückgespielt." -ForegroundColor Green
}

# =====================================================================
# FUNKTION 3: HTML-REPORT ERSTELLEN
# =====================================================================
function Create-HTMLReport {
    $reportPath = "$BackupRoot\Report_$timestamp.html"
    $results = Test-Certificates -Silent

    $style = @"
<style>
body { font-family: Segoe UI, sans-serif; margin: 20px; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 8px; }
th { background: #333; color: white; }
tr:nth-child(even) { background: #f2f2f2; }
.good { background-color: #c8e6c9; }
.warn { background-color: #fff9c4; }
.bad  { background-color: #ffcdd2; }
</style>
"@

    $rows = $results | ForEach-Object {
        $class = "good"
        if ($_.ValidSignature -eq $false) { $class = "bad" }
        elseif ($_.ChainStatus -ne "" -and $_.ChainStatus -ne "NoError") { $class = "warn" }

        "<tr class='$class'>
            <td>$($_.Store)</td>
            <td>$($_.Subject)</td>
            <td>$($_.Issuer)</td>
            <td>$($_.NotBefore)</td>
            <td>$($_.NotAfter)</td>
            <td>$($_.Thumbprint)</td>
            <td>$($_.ValidSignature)</td>
            <td>$($_.ChainStatus)</td>
        </tr>"
    }

    $html = @"
<html>
<head>
<title>Zertifikatsreport</title>
$style
</head>
<body>
<h2>Zertifikatsreport – $timestamp</h2>
<table>
<tr>
<th>Store</th>
<th>Subject</th>
<th>Issuer</th>
<th>NotBefore</th>
<th>NotAfter</th>
<th>Thumbprint</th>
<th>ValidSignature</th>
<th>ChainStatus</th>
</tr>
$rows
</table>
</body>
</html>
"@

    $html | Out-File -FilePath $reportPath -Encoding UTF8

    Write-Host "HTML-Report erstellt: $reportPath" -ForegroundColor Green
}

# =====================================================================
# FUNKTION 4: UNGÜLTIGE ZERTIFIKATE LÖSCHEN
# =====================================================================
function Remove-InvalidCertificates {
    $results = Test-Certificates -Silent

    $invalid = $results | Where-Object {
        $_.ValidSignature -eq $false -or
        ($_.ChainStatus -ne "" -and $_.ChainStatus -ne "NoError")
    }

    foreach ($item in $invalid) {
        Write-Host "Lösche ungültiges Zertifikat: $($item.Thumbprint)" -ForegroundColor Red
        Remove-Item $item.Cert.PSPath -Force
    }

    Write-Host "Alle ungültigen Zertifikate wurden gelöscht." -ForegroundColor Green
}

# =====================================================================
# FUNKTION 5: ECHTHEITSPRÜFUNG
# =====================================================================
function Test-Certificates {
    param([switch]$Silent)

    $results = @()

    foreach ($store in $Stores) {
        Get-ChildItem $store | ForEach-Object {
            $cert = $_

            $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
            $chain.ChainPolicy.RevocationMode = "Online"
            $chain.ChainPolicy.RevocationFlag = "EntireChain"
            $chain.ChainPolicy.VerificationFlags = "NoFlag"
            $chain.ChainPolicy.UrlRetrievalTimeout = (New-TimeSpan -Seconds 10)

            $isValid = $chain.Build($cert)
            $chainStatus = ($chain.ChainStatus | Select-Object -ExpandProperty Status -ErrorAction Ignore) -join ", "

            $results += [PSCustomObject]@{
                Store          = $store
                Subject        = $cert.Subject
                Issuer         = $cert.Issuer
                NotBefore      = $cert.NotBefore
                NotAfter       = $cert.NotAfter
                Thumbprint     = $cert.Thumbprint
                ValidSignature = $isValid
                ChainStatus    = $chainStatus
                Cert           = $cert
            }
        }
    }

    if (-not $Silent) {
        $results | Format-Table -AutoSize
    }

    return $results
}

# =====================================================================
# INTERAKTIVES MENÜ
# =====================================================================
function Show-Menu {
    Clear-Host
    Write-Host "================ Zertifikats-Tool ================" -ForegroundColor Cyan
    Write-Host "1) Backup erstellen"
    Write-Host "2) Backup zurückspielen"
    Write-Host "3) HTML-Report erstellen"
    Write-Host "4) Ungültige Zertifikate löschen"
    Write-Host "5) Echtheitsprüfung aller Zertifikate"
    Write-Host "6) Beenden"
    Write-Host "=================================================="
}

do {
    Show-Menu
    $choice = Read-Host "Auswahl"

    switch ($choice) {
        1 { Create-Backup }
        2 { Restore-Backup }
        3 { Create-HTMLReport }
        4 { Remove-InvalidCertificates }
        5 { Test-Certificates }
        6 { Write-Host "Beende Script." -ForegroundColor Yellow }
        default { Write-Host "Ungültige Auswahl." -ForegroundColor Red }
    }

    if ($choice -ne 6) {
        Write-Host ""
        Read-Host "Weiter mit ENTER"
    }

} while ($choice -ne 6)
