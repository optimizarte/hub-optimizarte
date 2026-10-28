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
    [string]$Token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJsb2dnZWQiOnRydWUsImlkIjoiNjk0IiwidXNlcklkIjo2NDYwLCJ0ZXNpc0lkIjo2Njc2LCJ1cHciOm51bGwsInZpc3VhbElkIjoiNjk0ICAiLCJjb2xJZCI6IjY5NCAgIiwiZW1haWwiOiJEQU5ZQE9QVElNSVpBUlRFLkNPTSAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAiLCJub21icmUiOiJIRVJOQU5ERVogU0VHVVJBLCBEQU5JRUwiLCJvcmdhbml6YWNpb24iOiJIRVJOQU5ERVogU0VHVVJBLCAgREFOWSAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIiwicGVyZmlsVXN1YXJpbyI6MiwiY29sYWJvcmFkb3JQcm8iOmZhbHNlLCJnZXN0aW9uSW5jaWRlbmNpYXNDbGllbnRlcyI6ZmFsc2UsImNhbWJpYXJQYXNzIjpmYWxzZSwicm9sZXMiOlsyXSwiY29uZmlndXJhY2lvbiI6eyJhc2lnbmFySW5jaWRlbmNpYXMiOmZhbHNlLCJhc2lnbmFyQ2FtYmlvTWVkaWFjaW9uIjpmYWxzZSwiYXNpZ25hclByb3llY3RvcyI6ZmFsc2UsImFzaWduYXJTaW5pZXN0cm9zIjpmYWxzZSwiYXNpZ25hclJlY2lib3MiOmZhbHNlLCJvY3VsdGFyQ29taXNpb25lc0NNIjpmYWxzZSwidHJhbWl0YVNpbmllc3Ryb3MiOmZhbHNlfSwiaWF0IjoiMTc5Mjg4NjQ1NSIsImV4cCI6MTc5MjkyOTY1NSwiaWRDb3JyZWR1cmlhIjo2MzksInBlcm1pc29BU00iOnsicHJlc3VwdWVzdG8iOnsibXNnIjoiRWwgdXN1YXJpbyBubyB0aWVuZSBwZXJtaXNvcyBlbiBUQk0gcGFyYSBBU00ifSwidGFyaWZpY2FjaW9uIjp7Im1zZyI6IkVsIHVzdWFyaW8gbm8gdGllbmUgcGVybWlzb3MgZW4gVEJNIHBhcmEgQVNNIn0sImNsaWVudGVBU00iOnsibXNnIjoiRWwgdXN1YXJpbyBubyB0aWVuZSBwZXJtaXNvcyBlbiBUQk0gcGFyYSBBU00ifX0sInVzdWFyaW9Bdmlzb3NWaXNpYmlsaWRhZCI6dHJ1ZSwibW9zdHJhckFjY2Vzb1RhYmxhc0l6cXVpZXJkYSI6ZmFsc2UsInJlY2liaXJNYWlsVHJhbWl0YWRvck9wb3J0dW5pZGFkIjpmYWxzZSwicmVjaWJpck1haWxUcmFtaXRhZG9yUmVjaWJvIjpmYWxzZSwiYXZpc29zTG9naW4iOmZhbHNlLCJ0dW5uZWwiOmZhbHNlLCJ1c2VySW1wIjpudWxsLCJleHBpcmVkUGFzc3dvcmQiOmZhbHNlLCJvZmljaW5hcyI6W10sImFjY2Vzb09maWNpbmEiOmZhbHNlLCJhc2lnbmFyT2ZpY2luYSI6ZmFsc2UsImFjY2Vzb1JlZENvbWVyY2lhbCI6ZmFsc2UsImFzaWduYXJDb2xhYm9yYWRvcmVzIjp0cnVlLCJjYW5TZW5kU21zIjpmYWxzZSwidGFyZWFzQ29sYWJvcmFkb3JlcyI6ZmFsc2V9.216CtcCuUxyiBwj7mkXpbCeSxo2hBhhnJmzlPtm1CcA",
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

# --- Descarregar polisses per obtenir companiaNombre (CIA) ---
Write-Host ""
Write-Host "Descarregant polisses (companiaNombre per client)..." -ForegroundColor Cyan

$polMap = @{}  # key: clienteId -> companiaNombre (última companyia guanyadora)
$tesisPol = @{}  # key: clienteId -> llista de polisses Tesis (id, num, mat, vto)
$polPage = 1
$polTotal = 0
do {
    $pUrl = "$BASE/listadoPolizas?grupos=poliza.lista,poliza.recibo&modoAsignacionColaborador=1&page=$polPage&limit=100"
    try {
        $pResp = Invoke-RestMethod -Uri $pUrl -Headers $headers -Method GET -ErrorAction Stop
    } catch {
        Write-Warning "Error descarregant polisses pagina $polPage`: $($_.Exception.Message.Split([char]10)[0])"
        break
    }
    $items = if ($pResp.PSObject.Properties["items"]) { $pResp.items } else { @() }
    if ($items.Count -eq 0) { break }
    foreach ($p in $items) {
        $cid = "$($p.clienteId)".Trim()
        $cia = "$($p.companiaNombre)".Trim()
        if ($cid -and $cia) { $polMap[$cid] = $cia }
        $tid = "$($p.id)".Trim()
        if ($cid -and $tid) {
            if (-not $tesisPol.ContainsKey($cid)) { $tesisPol[$cid] = New-Object System.Collections.ArrayList }
            $vtoRaw = "$($p.fechaVencimiento)"
            $vto = if ($vtoRaw.Length -ge 10) { $vtoRaw.Substring(0,10).Replace('-','') } else { '' }
            [void]$tesisPol[$cid].Add(@{ id = $tid; num = "$($p.numeroPoliza)".Trim(); mat = "$($p.matricula)".Trim(); vto = $vto })
        }
        $polTotal++
    }
    $polPage++
} while ($polPage -le 300)

Write-Host "Polisses processades: $polTotal  / clients amb CIA: $($polMap.Count)" -ForegroundColor Green

# Injectar CIA als clients descarregats
$ciaInjected = 0
foreach ($id in $allClients.Keys) {
    if ($polMap.ContainsKey($id) -and $polMap[$id]) {
        $allClients[$id]["cia"] = $polMap[$id]
        $ciaInjected++
    } else {
        $allClients[$id]["cia"] = ""
    }
}
Write-Host "CIA injectada a $ciaInjected clients" -ForegroundColor Green

# --- Actualitzar rawRows al seed de seguimiento-cartera-anagan.html ---
Write-Host ""
Write-Host "Actualitzant seed rawRows amb CIA..." -ForegroundColor Cyan

if (-not (Test-Path $HtmlPath)) {
    Write-Error "No es troba: $HtmlPath"
    exit 1
}

$htmlSeed = [System.IO.File]::ReadAllText($HtmlPath, [System.Text.Encoding]::UTF8)

# Extreure el bloc seed actual
$seedMatch = [System.Text.RegularExpressions.Regex]::Match(
    $htmlSeed,
    '(<script[^>]+id="seed"[^>]*>)([\s\S]*?)(</script>)'
)

if ($seedMatch.Success) {
    $seedContent = $seedMatch.Groups[2].Value
    $newSeedContent = $seedContent

    # Per cada clienteId, substituir CIA (a[1]) al rawRows del seed
    # El seed te files amb format: [clienteId,"",ef,...] (CIA sempre buida)
    # Substituim [clienteId,"", per [clienteId,"COMPANY",
    foreach ($cid in $polMap.Keys) {
        $cia = $polMap[$cid]
        if (-not $cia) { continue }
        $escapedCia = $cia.Replace('\','\\').Replace('"','\"')
        $oldToken = '["' + $cid + '","",'
        $newToken = '["' + $cid + '","' + $escapedCia + '",'
        if ($newSeedContent.IndexOf($oldToken) -ge 0) {
            $newSeedContent = $newSeedContent.Replace($oldToken, $newToken)
        }
    }

    if ($newSeedContent -ne $seedContent) {
        $htmlSeed = $htmlSeed.Replace($seedMatch.Groups[2].Value, $newSeedContent)
        [System.IO.File]::WriteAllText($HtmlPath, $htmlSeed, [System.Text.Encoding]::UTF8)
        Write-Host "rawRows actualitzats amb CIA." -ForegroundColor Green
    } else {
        Write-Host "rawRows: cap CIA nova per injectar." -ForegroundColor Yellow
    }
} else {
    Write-Warning "Bloc seed no trobat a $HtmlPath - CIA no injectada al seed."
}

# --- Serialitzar a JSON ---
$clientList = $allClients.Values | Sort-Object { [long]$_.id }

# Construir JSON manualment per evitar problemes d'encoding
$jsonLines = @()
foreach ($c in $clientList) {
    $line = [System.Text.StringBuilder]::new()
    $null = $line.Append("{")
    $null = $line.Append('"id":')
    $null = $line.Append((ConvertTo-Json $c.id -Compress))
    foreach ($k in @("nif","tomador","tel","email","dir","loc","prov","cp","tipDoc","fechaAlta","cia")) {
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

# Llegir de nou (pot haver estat actualitzat per la injecció CIA al seed)
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

# --- Bloc tesis-pol: id de polissa Tesis per obrir la fitxa a ANAGAN des del CRM ---
function JsEsc([string]$s) { return $s.Replace('\','\\').Replace('"','\"').Replace('<','\u003c') }
$tpParts = New-Object System.Collections.ArrayList
foreach ($cid in $tesisPol.Keys) {
    $tpItems = @()
    foreach ($t in $tesisPol[$cid]) {
        $tpItems += '{"id":"' + (JsEsc $t.id) + '","num":"' + (JsEsc $t.num) + '","mat":"' + (JsEsc $t.mat) + '","vto":"' + (JsEsc $t.vto) + '"}'
    }
    [void]$tpParts.Add('"' + (JsEsc $cid) + '":[' + ($tpItems -join ',') + ']')
}
$tpBlock = '<script type="application/json" id="tesis-pol" data-ts="' + $ts + '">{' + ($tpParts -join ',') + '}</script>'
$tpStart = $html.IndexOf('<script type="application/json" id="tesis-pol"')
if ($tpStart -ge 0) {
    $tpEnd = $html.IndexOf('</script>', $tpStart) + 9
    $html = $html.Substring(0, $tpStart) + $tpBlock + $html.Substring($tpEnd)
} else {
    $csStart = $html.IndexOf('id="clients-seed"')
    $csEnd = $html.IndexOf('</script>', $csStart) + 9
    $html = $html.Substring(0, $csEnd) + "`n" + $tpBlock + $html.Substring($csEnd)
}
Write-Host "Bloc tesis-pol: $($tesisPol.Count) clients amb ID de polissa Tesis." -ForegroundColor Green

[System.IO.File]::WriteAllText($HtmlPath, $html, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "DONE! $($allClients.Count) clients injectats a l'app." -ForegroundColor Green
Write-Host "Refresca el navegador per veure les dades." -ForegroundColor Cyan
Write-Host ""

# Copiar a l'extensió OPTICRM perquè ANAGAN sigui accessible des del CRM
$ExtPath = "C:\Users\primo\OneDrive - OPTIMIZARTE 3.0 - SCO\!!IA\!!!PortalOccident\!!OPTICRM\extension\seguimiento-cartera-anagan.html"
try {
    Copy-Item $HtmlPath $ExtPath -Force
    Write-Host "Copiat a l'extensió OPTICRM: $ExtPath" -ForegroundColor Green
} catch {
    Write-Host "AVIS: no s'ha pogut copiar a l'extensió: $_" -ForegroundColor Yellow
}
# Sincronitzar la resta d'apps del HUB amb l'extensio (ENPA seguiment mensual, etc.)
$SyncScript = Join-Path $PSScriptRoot "Sincronitzar-Apps-OPTICRM.ps1"
if (Test-Path $SyncScript) { & $SyncScript }
Write-Host ""
