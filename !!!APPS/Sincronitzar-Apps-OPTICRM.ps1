# Sincronitzar-Apps-OPTICRM.ps1
# Copia les apps HTML del HUB a l'extensio OPTICRM (nomes si han canviat).
# L'extensio les carrega com a blob dins el layout del CRM: despres de copiar cal
# recarregar l'extensio (chrome://extensions) i la pestanya del CRM.
# Nota: fitxer ASCII (PowerShell 5.1).

$Apps = "C:\Users\primo\OneDrive - OPTIMIZARTE 3.0 - SCO\!!IA\!!!PortalOccident\!!GitHub-Repository-HUB\HUB-OPTIMIZARTE\!!!APPS"
$Ext  = "C:\Users\primo\OneDrive - OPTIMIZARTE 3.0 - SCO\!!IA\!!!PortalOccident\!!OPTICRM\extension"

# origen (HUB) -> desti (extensio)
$Map = @(
    @{ Src = "seguiment-mensual_3.html";        Dst = "seguiment-mensual.html" },
    @{ Src = "seguimiento-cartera-anagan.html"; Dst = "seguimiento-cartera-anagan.html" }
)

$changed = 0
foreach ($m in $Map) {
    $src = Join-Path $Apps $m.Src
    $dst = Join-Path $Ext  $m.Dst
    if (-not (Test-Path $src)) { Write-Host ("FALTA origen: " + $src) -ForegroundColor Yellow; continue }
    $same = $false
    if (Test-Path $dst) {
        $same = ((Get-FileHash $src -Algorithm MD5).Hash -eq (Get-FileHash $dst -Algorithm MD5).Hash)
    }
    if ($same) {
        Write-Host ("OK (sense canvis): " + $m.Dst) -ForegroundColor DarkGray
    } else {
        try {
            Copy-Item $src $dst -Force
            Write-Host ("COPIAT: " + $m.Src + " -> extension\" + $m.Dst) -ForegroundColor Green
            $changed++
        } catch {
            Write-Host ("ERROR copiant " + $m.Src + ": " + $_) -ForegroundColor Red
        }
    }
}

if ($changed -gt 0) {
    Write-Host ""
    Write-Host "$changed fitxer(s) actualitzat(s). Recarrega l'extensio i la pestanya del CRM." -ForegroundColor Cyan
} else {
    Write-Host "Tot sincronitzat." -ForegroundColor Cyan
}
