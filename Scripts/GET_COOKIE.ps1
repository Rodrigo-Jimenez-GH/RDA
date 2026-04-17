param(
    [string]$SavePath = "C:\Macros\RDA\FGC\COOKIE.txt",
    [string]$EXPIRY = "C:\Macros\RDA\FGC\EXPIRY.txt",
    [bool]$CloseEdge = $false
)

$logPath = "$env:TEMP\FGC_Cookie.log"
Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue

# --- Configuration ---
$TargetURL = "https://myapps-atl01.secure.fedex.com/clearance/manifest/"
$edge = "msedge.exe"
$tempProfile = "$env:TEMP\edge_debug_profile"
$global:CDPCommandId = 100

# --- Función Send-CDPCommand (Exact replication) ---
function Send-CDPCommand {
    param(
        [Parameter(Mandatory=$true)]
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory=$true)]
        [hashtable]$Params
    )
    $id = $global:CDPCommandId
    $global:CDPCommandId++
    $json = @{ id = $id; method = $Params.method; params = $Params.params } | ConvertTo-Json -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $segment = New-Object System.ArraySegment[byte] -ArgumentList (, $bytes)
    $Socket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).Wait()
    $buffer = New-Object byte[] 102400 
    $result = $Socket.ReceiveAsync([ArraySegment[byte]]$buffer, [Threading.CancellationToken]::None).Result
    return [Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count) | ConvertFrom-Json
}

# --- Función Wait-TabReady (Exact replication) ---
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
    while (-not $ready) {
        try {
            $urlResp = Send-CDPCommand $Socket @{ method = "Runtime.evaluate"; params = @{ expression = "window.location.href" } }
            $stateResp = Send-CDPCommand $Socket @{ method = "Runtime.evaluate"; params = @{ expression = "document.readyState" } }
            $currentUrl = if ($urlResp.result -and $urlResp.result.result) { $urlResp.result.result.value } else { $null }
            $state = if ($stateResp.result -and $stateResp.result.result) { $stateResp.result.result.value } else { $null }
            $urlOk = if ($ExpectedUrl) { $currentUrl -eq $ExpectedUrl } else { $true }
            if ($urlOk -and $state -eq "complete") { $ready = $true }
            else { Start-Sleep -Milliseconds $RetryDelayMs }
        } catch { throw "Error en Wait-TabReady: $_" }
    }
    return @{ url = $currentUrl; state = $state }
}

# --- Bloque principal ---
try {
    Write-Host "--- Script Start: $(Get-Date) ---" -ForegroundColor Cyan
    Write-Host "Iniciando Edge..."
    Start-Process $edge "--remote-debugging-port=9222 --user-data-dir=`"$tempProfile`" $TargetURL"

    $tab = $null
    while (-not $tab) {
        try {
            $tabs = Invoke-RestMethod http://localhost:9222/json
            $tab = $tabs | Where-Object { $_.url -like "$TargetURL*" } | Select-Object -First 1
        } catch { Start-Sleep -Milliseconds 500 }
    }

    $socket = New-Object System.Net.WebSockets.ClientWebSocket
    $socket.ConnectAsync([Uri]$tab.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()

    # 1. Wait for document.readyState
    Wait-TabReady -Socket $socket -ExpectedUrl "$TargetURL"

    # 2. Validation Loop for JSESSIONID
    Write-Host "Esperando validación de sesión (Buscando JSESSIONID)..." -ForegroundColor Yellow
    $authenticated = $false
    $cookies = $null

    while (-not $authenticated) {
        $cookieResp = Send-CDPCommand $socket @{ method = "Network.getCookies"; params = @{} }
        $cookies = $cookieResp.result.cookies
        
        if ($cookies | Where-Object { $_.name -eq "JSESSIONID" }) {
            $authenticated = $true
            Write-Host "JSESSIONID detectada. Sesión lista." -ForegroundColor Green
        } else {
            Write-Host "JSESSIONID no encontrada aún. Reintentando..."
            Start-Sleep -Milliseconds 1500
        }
    }

    # --- INPECCIÓN DE COOKIES (DEBUG) ---
    Write-Host "`n### INSPECCIÓN DE TODAS LAS COOKIES DETECTADAS ###" -ForegroundColor Gray
    $cookies | ForEach-Object {
        $expDate = "Session"
        if ($_.expires -gt 0) {
            $expDate = (Get-Date "1970-01-01").AddSeconds($_.expires).ToLocalTime().ToString("MM/dd/yyyy HH:mm:ss")
        }
        
        Write-Host "--------------------------------------------------"
        Write-Host "NOMBRE:  $($_.name)" -ForegroundColor White
        Write-Host "VALOR:   $($_.value)" -ForegroundColor Gray
        Write-Host "DOMINIO: $($_.domain)"
        Write-Host "EXPIRA:  $expDate" -ForegroundColor Cyan
    }
    Write-Host "--------------------------------------------------`n"

    # 3. Process and Save
    $allCookiesString = ($cookies | ForEach-Object { "$($_.name)=$($_.value)" }) -join ";"
    
    # Expiry logic: 1-hour ceiling
    $now = Get-Date
    $persistent = $cookies | Where-Object { $_.expires -gt 0 }
    if ($persistent) {
        $earliest = ($persistent | Measure-Object -Property expires -Minimum).Minimum
        $calculated = (Get-Date "1970-01-01").AddSeconds($earliest).ToLocalTime()
        $finalExpiry = if ($calculated -gt $now.AddHours(1)) { $now.AddHours(1) } else { $calculated }
    } else {
        $finalExpiry = $now.AddHours(1)
    }

    $utf8NoBOM = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($SavePath, $allCookiesString, $utf8NoBOM)
    [System.IO.File]::WriteAllText($EXPIRY, $finalExpiry.ToString("MM/dd/yyyy HH:mm:ss"), $utf8NoBOM)

    Write-Host "Success! Cookies y Expiración guardados." -ForegroundColor Green

} catch {
    Write-Host "Error detectado: $_" -ForegroundColor Red
    $CloseEdge = $true # Force close on error to reset
} finally {
    if ($socket -and $socket.State -eq 'Open') {
        $socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "", [Threading.CancellationToken]::None).Wait()
    }

    # CONDITIONAL CLOSE
    if ($CloseEdge) {
        Write-Host "Cerrando Edge por configuración..." -ForegroundColor Yellow
        Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*--remote-debugging-port=9222*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    } else {
        Write-Host "Edge se mantiene abierto para inspección manual." -ForegroundColor Magenta
    }

    Stop-Transcript | Out-Null
}