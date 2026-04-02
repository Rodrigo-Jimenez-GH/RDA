param(
    [string]$TargetURL,
    [string]$path,
    [string]$pathExpiry
)


if (-not (Test-Path $pathExpiry)) {
    Write-Host "File missing → expired"
    New-Item -Path $pathExpiry -ItemType File -Force | Out-Null
    "01/01/2000 00:00:00" | Set-Content $pathExpiry
}
try {
    $content = Get-Content $pathExpiry -Raw
    $timestamp = [datetime]::ParseExact($content.Trim(), "MM/dd/yyyy HH:mm:ss", $null)
    $now = Get-Date
    if ($timestamp -ge $now) {exit}
} catch {
    Write-Host "Invalid file content → treating as expired"
}

$edge = "msedge.exe"
$tempProfile = "$env:TEMP\edge_debug_profile"
$global:CDPCommandId = 100

# --- Función Send-CDPCommand ---
function Send-CDPCommand {
    param(
        [Parameter(Mandatory=$true)]
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Params
    )

    $id = $global:CDPCommandId
    $global:CDPCommandId++

    $json = @{
        id = $id
        method = $Params.method
        params = $Params.params
    } | ConvertTo-Json -Compress

    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $segment = New-Object System.ArraySegment[byte] -ArgumentList (, $bytes)
    $Socket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).Wait()

    $buffer = New-Object byte[] 10240
    $result = $Socket.ReceiveAsync([ArraySegment[byte]]$buffer, [Threading.CancellationToken]::None).Result

    return [Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count) | ConvertFrom-Json
}

# --- Función Wait-TabReady ---
function Wait-TabReady {
    param(
        [Parameter(Mandatory=$true)]
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory=$false)]
        [string]$ExpectedUrl = $null,
        [Parameter(Mandatory=$false)]
        [int]$RetryDelayMs = 3000
    )

    $ready = $false
    $attempt = 0

    while (-not $ready) {
        try {
            $urlResp = Send-CDPCommand $Socket @{ method = "Runtime.evaluate"; params = @{ expression = "window.location.href" } }
            $stateResp = Send-CDPCommand $Socket @{ method = "Runtime.evaluate"; params = @{ expression = "document.readyState" } }

            $currentUrl = if ($urlResp.result -and $urlResp.result.result) { $urlResp.result.result.value } else { $null }
            $state = if ($stateResp.result -and $stateResp.result.result) { $stateResp.result.result.value } else { $null }

            $urlOk = $true
            if ($ExpectedUrl) { $urlOk = $currentUrl -eq $ExpectedUrl }

            if ($urlOk -and $state -eq "complete") { $ready = $true }
            else { Start-Sleep -Milliseconds $RetryDelayMs; $attempt++; Write-Host "Esperando carga... intento $attempt ($currentUrl | $state)" }
        }
        catch {
            throw "Error en Wait-TabReady: $_"
        }
    }

    Write-Host "Página cargada: $currentUrl"
    return @{ url = $currentUrl; state = $state }
}

# --- Bloque principal con manejo de errores ---
try {
    # Abrir Edge
    $proc = Start-Process $edge "--remote-debugging-port=9222 --user-data-dir=`"$tempProfile`" $TargetURL"

    # Esperar pestaña RDA
    $tab = $null
    $X = 0
    while (-not $tab) {
        try {
            $tabs = Invoke-RestMethod http://localhost:9222/json
            Write-Host "Tabs count:" $tabs.Count

            $tab = $tabs | Where-Object { $_.url -like "$TargetURL*" } | Select-Object -First 1
            if ($tab) { Write-Host "FEDEXAPP tab found:" $tab.url } else { Write-Host "FEDEXAPP tab not found yet..." }
        }
        catch { 
            Write-Host "DevTools not ready..."
        }

        Start-Sleep -Milliseconds 500
        $X++
        Write-Host "Retry: $X"
    }

    # Conectar WebSocket
    $ws = $tab.webSocketDebuggerUrl
    $socket = New-Object System.Net.WebSockets.ClientWebSocket
    $uri = [Uri]$ws
    $socket.ConnectAsync($uri, [Threading.CancellationToken]::None).Wait()

    # Esperar página cargada
    Wait-TabReady -Socket $socket -ExpectedUrl "$TargetURL"

    # Esperar token en localStorage
    $tokenReady = $false
    while (-not $tokenReady) {
        $resp = Send-CDPCommand $socket @{
            method = "Runtime.evaluate"
            params = @{ expression = 'localStorage.getItem("okta-token-storage") !== null' }
        }
        if ($resp.result -and $resp.result.result) { $tokenReady = $resp.result.result.value }
        if (-not $tokenReady) { Start-Sleep -Milliseconds 500 }
    }

    # Esperar token válido
    $validToken = $false
    while (-not $validToken) {
        $resp = Send-CDPCommand $socket @{
            method = "Runtime.evaluate"
            params = @{ expression = 'JSON.stringify(JSON.parse(localStorage.getItem("okta-token-storage")).accessToken)' }
        }

        if (-not $resp.result -or -not $resp.result.result -or -not $resp.result.result.value) { Start-Sleep 0.3; continue }

        $data = $resp.result.result.value | ConvertFrom-Json
        $token = $data.accessToken
        $expires = $data.expiresAt

        if (-not $token) { Write-Host "Token aún no disponible..."; Start-Sleep 1; continue }

        $expiry = [DateTimeOffset]::FromUnixTimeSeconds($expires).ToLocalTime().DateTime
        if ($expiry -le (Get-Date)) {
            Write-Host "TOKEN EXPIRADO EL: $expiry"
            Write-Host "TOKEN EXPIRADO: $token"
            Send-CDPCommand $socket @{ method="Page.reload"; params=@{} }
            Wait-TabReady -Socket $socket -ExpectedUrl "$TargetURL"
            Write-Host "Recarga Lista"
            continue
        }

        $validToken = $true
    }

    # Limpiar BOM y guardar
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($token)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $bytes = $bytes[3..($bytes.Length - 1)]
    }
    $tokenClean = [System.Text.Encoding]::UTF8.GetString($bytes)

$utf8NoBOM = New-Object System.Text.UTF8Encoding($false)

[System.IO.File]::WriteAllText($path, $tokenClean, $utf8NoBOM)
[System.IO.File]::WriteAllText($pathExpiry, $expiry, $utf8NoBOM)

    Write-Host "EXPIRA: $expiry"
    Write-Host "TOKEN: $token"
    Write-Host "TOKEN GUARDADO en $path"

    # Cerrar Edge
    Get-CimInstance Win32_Process |
        Where-Object { $_.CommandLine -like '*--remote-debugging-port=9222*' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force }

    Write-Host "EDGE CERRADO"
}
catch {
    Write-Host "¡ERROR DETECTADO! Detalles:"
    Write-Host $_
    
    # Intentar cerrar Edge si sigue abierto
    try {
        Get-CimInstance Win32_Process |
            Where-Object { $_.CommandLine -like '*--remote-debugging-port=9222*' } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    } catch {}
    pause
    Exit 1
}
