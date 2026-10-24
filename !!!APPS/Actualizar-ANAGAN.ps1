# =============================================================
#  Actualizar-ANAGAN.ps1
#  Descarrega les dades de clients de l'API TESIS i les injecta
#  a seguimiento-cartera-anagan.html (bloc <script id="clients-seed">)
#
#  COM OBTENIR EL TOKEN:
#  1. Ves a https://anagan.app.tesis.io amb Edge (sessió activa)
#  2. F12 > Network > Fetch/XHR
#  3. Refresca la pàgina
#  4. Clica qualsevol petició a api.tesis.io > Headers > Request Headers
#  5. Copia el valor de "Authorization" (Bearer xxxxx...)
# =============================================================

param(
    [string]$Token = "",
    [string]$HtmlPath = "$PSScriptRoot\seguimiento-cartera-anagan.html"
)

# --- Token ---
if (-not $Token) {
    Write-Host ""
    Write-Host "TOKEN BEARER" -ForegroundColor Cyan
    Write-Host "Pega el valor del header Authorization de DevTools (incluye 'Bearer ...'):" -ForegroundColor Yellow
    $Token = Read-Host "Authorization"
}
if (-not $Token) { Write-Error "Token requerit."; exit 1 }
if (-not $Token.StartsWith("Bearer ")) { $Token = "Bearer $Token" }

# --- Config ---
$BASE = "https://api.tesis.io/v2"
$GRUPOS_CLIENTS = "cliente.lista,cliente.direccionPrincipal,cliente.clienteAsurnet,cliente.colaborador1"
$LIMIT = 50

$headers = @{
    "Authorization" = $Token
    "Accept"        = "application/json"
    "Content-Type"  = "application/json"
    "Origin"        = "https://anagan.app.tesis.io"
    "Referer"       = "https://anagan.app.tesis.io/"
}

Write-Host ""
Write-Host "Connectant amb l'API TESIS..." -ForegroundColor Cyan

# --- Obtenir total de clients ---
try {
    $countUrl = "$BASE/countClientes?grupos=$GRUPOS_CLIENTS&modoAsignacionColaborador=1&estadoSiniestros=todos&informacionComercial=false"
    $countRes = Invoke-RestMethod -Uri $countUrl -Headers $headers -Method GET
    $total = $countRes.total
    $totalPages = [Math]::Ceiling($total / $LIMIT)
    Write-Host "Total clients: $total  (pagines: $totalPages)" -ForegroundColor Green
} catch {
    $connErr = $_.Exception.Message; Write-Error "Error connectant a l'API: $connErr"
    Write-Host "Comprova que el token no ha caducat i torna-ho a intentar." -ForegroundColor Yellow
    exit 1
}

# --- Descarregar totes les pagines ---
$allClients = @{}
$progress = 0

for ($page = 1; $page -le $totalPages; $page++) {
    $pct = [int](($progress / $total) * 100)
    Write-Progress -Activity "Descarregant clients ANAGAN" -Status "Pagina $page de $totalPages ($progress/$total)" -PercentComplete $pct

    $url = "$BASE/listadoClientes?grupos=$GRUPOS_CLIENTS&page=$page&limit=$LIMIT&modoAsignacionColaborador=1&estadoSiniestros=todos&informacionComercial=false&sort=cliente.fechaAlta&inverse=true"

    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method GET
        $data = if ($resp.PSObject.Properties["items"]) { $resp.items } else { $resp }
    } catch {
        $errMsg = $_.Exception.Message; Write-Warning "Error pagina $page`: $errMsg. Reintentant..."
        Start-Sleep -Seconds 2
        try {
            $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method GET
            $data = if ($resp.PSObject.Properties["items"]) { $resp.items } else { $resp }
        } catch {
            $errMsg2 = $_.Exception.Message; Write-Error "Error irrecuperable pagina $page`: $errMsg2"
            exit 1
        }
    }

    foreach ($c in $data) {
        $id = "$($c.id)".Trim()
        if (-not $id) { continue }

        # Intentar trobar el NIF real (pot estar en nif, cif, o numeroDocumento)
        $nif = ""
        if ($c.PSObject.Properties["nif"])      { $nif = "$($c.nif)".Trim() }
        if (-not $nif -and $c.PSObject.Properties["cif"])  { $nif = "$($c.cif)".Trim() }
        if (-not $nif -and $c.PSObject.Properties["numeroDocumento"]) { $nif = "$($c.numeroDocumento)".Trim() }

        # Telefons
        $tel = ""
        foreach ($f in @("telefonoMovil","telefonoPrincipal","telefonoMovil2","telefonoPrincipal2")) {
            if ($c.PSObject.Properties[$f] -and "$($c.$f)".Trim()) { $tel = "$($c.$f)".Trim(); break }
        }

        # Email
        $email = ""
        if ($c.PSObject.Properties["email"]) { $email = "$($c.email)".Trim() }
        if (-not $email -and $c.PSObject.Properties["emailPrincipal"]) { $email = "$($c.emailPrincipal)".Trim() }

        $allClients[$id] = @{
            id        = $id
            nif       = $nif
            tomador   = "$($c.tomador)".Trim()
            tel       = $tel
            email     = $email
            dir       = "$($c.direccionCompletaFiscal)".Trim()
            loc       = "$($c.localidad)".Trim()
            prov      = "$($c.provincia)".Trim()
            cp        = "$($c.cp)".Trim()
            tipDoc    = "$($c.tipoDocumento)".Trim()
            fechaAlta = "$($c.fechaAlta)".Trim()
        }
        $progress++
    }

    Start-Sleep -Milliseconds 100  # respecte a l'API
}

Write-Progress -Activity "Descarregant clients ANAGAN" -Completed
Write-Host "Clients descarregats: $($allClients.Count)" -ForegroundColor Green

# --- Serialitzar a JSON ---
$clientList = $allClients.Values | Sort-Object { [long]$_.id }

# Construir JSON manualment per evitar problemes d'encoding
$jsonLines = @()
foreach ($c in $clientList) {
    $line = [System.Text.StringBuilder]::new()
    $null = $line.Append("{")
    $null = $line.Append('"id":')
    $null = $line.Append((ConvertTo-Json $c.id -Compress))
    foreach ($k in @("nif","tomador","tel","email","dir","loc","prov","cp","tipDoc","fechaAlta")) {
        $null = $line.Append(',"')
        $null = $line.Append($k)
        $null = $line.Append('":')
        $null = $line.Append((ConvertTo-Json ($c[$k] -replace '[\x00-\x1f]','') -Compress))
    }
    $null = $line.Append("}")
    $jsonLines += $line.ToString()
}
$jsonPayload = "[" + ($jsonLines -join ",") + "]"

# --- Injectar a l'HTML ---
if (-not (Test-Path $HtmlPath)) {
    Write-Error "No es troba: $HtmlPath"
    exit 1
}

Write-Host "Actualitzant HTML: $HtmlPath" -ForegroundColor Cyan

$html = [System.IO.File]::ReadAllText($HtmlPath, [System.Text.Encoding]::UTF8)

$ts = (Get-Date).ToString("yyyy-MM-dd HH:mm")
$newBlock = "<script type=`"application/json`" id=`"clients-seed`" data-ts=`"$ts`">$jsonPayload</script>"

# Si ja existeix el bloc, el reemplacem
if ($html -match '<script[^>]+id="clients-seed"[^>]*>') {
    $html = [System.Text.RegularExpressions.Regex]::Replace(
        $html,
        '<script[^>]+id="clients-seed"[^>]*>[\s\S]*?</script>',
        $newBlock
    )
    Write-Host "Bloc clients-seed actualitzat." -ForegroundColor Green
} else {
    # Inserir despres del bloc seed existent
    $html = $html -replace '(<script[^>]+id="seed"[^>]*>[\s\S]*?</script>)', "`$1`n$newBlock"
    Write-Host "Bloc clients-seed inserit." -ForegroundColor Green
}

[System.IO.File]::WriteAllText($HtmlPath, $html, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "DONE! $($allClients.Count) clients injectats a l'app." -ForegroundColor Green
Write-Host "Refresca el navegador per veure les dades." -ForegroundColor Cyan
Write-Host ""
