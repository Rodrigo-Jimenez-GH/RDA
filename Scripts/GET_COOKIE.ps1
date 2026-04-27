param(
    [string]$SavePath,
    [string]$EXPIRY,
    [bool]$CloseEdge = $false
)

$logPath = "$env:TEMP\FGC_Cookie.log"
Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue

# --- Configuration ---
#$TargetURL = "https://myapps-atl01.secure.fedex.com/clearance/manifest/"
$TargetURL = "https://fgc-lac-cairo-atl.prod.cloud.fedex.com/clearance/mainMenu.jsp"
$edge = "msedge.exe"
$tempProfile = "$env:TEMP\edge_debug_profile"
$global:CDPCommandId = 100

# --- WinAPI ---
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@

# --- DEBUG + HIDE FUNCTION ---
function Debug-Hide-EdgeWindow {

    Write-Host "`n=== DEBUG EDGE WINDOWS ===" -ForegroundColor Cyan

    $procs = Get-CimInstance Win32_Process |
        Where-Object { $_.Name -eq "msedge.exe" }

    foreach ($proc in $procs) {
        try {
            $p = Get-Process -Id $proc.ProcessId -ErrorAction Stop
            $hasWindow = $p.MainWindowHandle -ne 0

            Write-Host "----------------------------------------"
            Write-Host "PID:        $($p.Id)" -ForegroundColor Yellow
            Write-Host "HasWindow:  $hasWindow"
            Write-Host "Title:      $($p.MainWindowTitle)"
            Write-Host "CmdLine:    $($proc.CommandLine)"

            # 🎯 ONLY hide debugger Edge
            if ($hasWindow -and $proc.CommandLine -like "*--remote-debugging-port=9222*") {

                for ($i = 0; $i -lt 3; $i++) {
                    $success = [WinAPI]::ShowWindowAsync($p.MainWindowHandle, 0)
                    Start-Sleep -Milliseconds 200
                }

                Write-Host ">>> TARGET HIDDEN <<<" -ForegroundColor Green
                Write-Host "Result: $success"
            }

        } catch {
            Write-Host "Error con PID $($proc.ProcessId)" -ForegroundColor Red
        }
    }

    Write-Host "========================================`n"
}

# --- CDP COMMAND ---
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

# --- WAIT TAB READY ---
function Wait-TabReady {
    param(
        [Parameter(Mandatory=$true)]
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [string]$ExpectedUrl = $null,
        [int]$RetryDelayMs = 3000
    )

    while ($true) {
        $urlResp = Send-CDPCommand $Socket @{ method = "Runtime.evaluate"; params = @{ expression = "window.location.href" } }
        $stateResp = Send-CDPCommand $Socket @{ method = "Runtime.evaluate"; params = @{ expression = "document.readyState" } }

        $currentUrl = $urlResp.result.result.value
        $state = $stateResp.result.result.value

        write-host $currentUrl
        write-host $ExpectedUrl
        write-host "Current: $state"
        write-host (($ExpectedUrl -eq $null -or $currentUrl -eq $ExpectedUrl) -and $state -eq "complete")

        if (($ExpectedUrl -eq $null -or $currentUrl -eq $ExpectedUrl) -and $state -eq "complete") {
            return
        }

        Start-Sleep -Milliseconds $RetryDelayMs
    }
}

# --- MAIN ---
try {
    Write-Host "--- Script Start: $(Get-Date) ---" -ForegroundColor Cyan
    Write-Host "Iniciando Edge..."

    $edgeProcess = Start-Process $edge `
        "--remote-debugging-port=9222 --user-data-dir=`"$tempProfile`" $TargetURL" `
        -PassThru

    # Wait for debugger tab
    do {
        Start-Sleep -Milliseconds 500
        try {
            $tabs = Invoke-RestMethod http://localhost:9222/json
            $tab = $tabs | Where-Object { $_.url -like "$TargetURL*" } | Select-Object -First 1
        } catch {}
    } while (-not $tab)

    $socket = New-Object System.Net.WebSockets.ClientWebSocket
    $socket.ConnectAsync([Uri]$tab.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()

    Wait-TabReady -Socket $socket -ExpectedUrl $TargetURL

    # Wait for JSESSIONID
    Write-Host "Esperando validación de sesión..." -ForegroundColor Yellow

    do {
        $cookieResp = Send-CDPCommand $socket @{ method = "Network.getCookies"; params = @{} }
        $cookies = $cookieResp.result.cookies
        $found = $cookies | Where-Object { $_.name -eq "JSESSIONID" }

        if (-not $found) {
            Start-Sleep -Milliseconds 1500
        }

    } while (-not $found)

    Write-Host "Sesión lista." -ForegroundColor Green

    # Save cookies
    $cookieString = ($cookies | ForEach-Object { "$($_.name)=$($_.value)" }) -join ";"

    $now = Get-Date
    $finalExpiry = $now.AddHours(1)

    $utf8NoBOM = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($SavePath, $cookieString, $utf8NoBOM)
    [System.IO.File]::WriteAllText($EXPIRY, $finalExpiry.ToString("MM/dd/yyyy HH:mm:ss"), $utf8NoBOM)

    Write-Host "Cookies guardadas." -ForegroundColor Green

} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    $CloseEdge = $true
}
finally {
    if ($socket -and $socket.State -eq 'Open') {
        $socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "", [Threading.CancellationToken]::None).Wait()
    }

    if ($CloseEdge) {
        Get-CimInstance Win32_Process |
            Where-Object { $_.CommandLine -like '*--remote-debugging-port=9222*' } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    }

    Stop-Transcript | Out-Null

    Start-Sleep -Seconds 2
    Debug-Hide-EdgeWindow
}