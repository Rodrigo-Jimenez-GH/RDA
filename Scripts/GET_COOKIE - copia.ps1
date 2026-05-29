param(
    [string]$SavePath = "C:\Macros\RESPS\TOKEN.txt",
    [string]$EXPIRY = "C:\Macros\RESPS\TOKEN_EOPS.txt",
    [bool]$CloseEdge = $false
)

$logPath = "$env:TEMP\FGC_Cookie.log"
Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue

$TargetURL = "https://fgc-gui-app.app.paas.fedex.com/#/portal"
$PortalURL = "https://fgc-gui-app.app.paas.fedex.com/#/portal"
$edge = "msedge.exe"
$tempProfile = "$env:TEMP\edge_debug_profile"
$global:CDPCommandId = 100
$global:DevToolsPort = 9222

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@

function Debug-Hide-EdgeWindow {
    $procs = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq "msedge.exe" }

    foreach ($proc in $procs) {
        try {
            $p = Get-Process -Id $proc.ProcessId -ErrorAction Stop
            if ($p.MainWindowHandle -ne 0 -and $proc.CommandLine -like "*--remote-debugging-port=$global:DevToolsPort*") {
                for ($i = 0; $i -lt 3; $i++) {
                    [WinAPI]::ShowWindowAsync($p.MainWindowHandle, 0)
                    Start-Sleep -Milliseconds 200
                }
            }
        } catch {}
    }
}

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

    $Socket.SendAsync(
        $segment,
        [System.Net.WebSockets.WebSocketMessageType]::Text,
        $true,
        [Threading.CancellationToken]::None
    ).Wait()

    $buffer = New-Object byte[] 102400
    $result = $Socket.ReceiveAsync(
        [ArraySegment[byte]]$buffer,
        [Threading.CancellationToken]::None
    ).Result

    $text = [Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count)
    $obj = $text | ConvertFrom-Json

    if ($obj -is [System.Array]) { return $obj[-1] }
    return $obj
}

function Wait-TabReady {
    param(
        [Parameter(Mandatory=$true)]
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [string]$ExpectedUrl = $null,
        [int]$RetryDelayMs = 2000
    )

    while ($true) {
        $urlResp = Send-CDPCommand $Socket @{
            method = "Runtime.evaluate"
            params = @{ expression = "window.location.href" }
        }

        $stateResp = Send-CDPCommand $Socket @{
            method = "Runtime.evaluate"
            params = @{ expression = "document.readyState" }
        }

        $currentUrl = $urlResp.result.result.value
        $state = $stateResp.result.result.value

        Write-Host "Attached to: $currentUrl"
        Write-Host "Expected:    $ExpectedUrl"
        Write-Host "State:       $state"

        if (($ExpectedUrl -eq $null -or $currentUrl -like "$ExpectedUrl*") -and $state -eq "complete") {
            return $currentUrl
        }

        Start-Sleep -Milliseconds $RetryDelayMs
    }
}

function Close-OtherTabs {
    param(
        [string]$KeepTabId,
        [int]$Port = $global:DevToolsPort
    )

    try {
        $tabs = Invoke-RestMethod -Uri "http://localhost:$Port/json" -ErrorAction Stop

        foreach ($t in $tabs) {
            if (-not $t.id -or $t.id -eq $KeepTabId) { continue }

            try {
                Invoke-RestMethod -Uri "http://localhost:$Port/json/close/$($t.id)" -Method GET -TimeoutSec 5 -ErrorAction Stop | Out-Null
                Write-Host "Closed tab: $($t.url)"
            } catch {
                Write-Host "Failed closing tab: $($t.url)"
            }
        }
    } catch {
        Write-Host "Close-OtherTabs failed: $_"
    }
}

function Wait-ForTargetTab {
    param(
        [string]$TargetUrl,
        [int]$Port = $global:DevToolsPort,
        [int]$MaxWaitSec = 300,
        [int]$RetryDelayMs = 2000
    )

    $sw = [Diagnostics.Stopwatch]::StartNew()

    while ($sw.Elapsed.TotalSeconds -lt $MaxWaitSec) {
        try {
            $tabs = Invoke-RestMethod -Uri "http://localhost:$Port/json" -ErrorAction Stop

            Write-Host "`n--- DevTools targets ---"
            $tabs | ForEach-Object {
                Write-Host ("id={0} url={1} ws={2}" -f $_.id, $_.url, ($_.webSocketDebuggerUrl -as [string]))
            }
            Write-Host "------------------------`n"

            $tab = $tabs |
                Where-Object {
                    $_.url -and ($_.url -like "$TargetUrl*")
                } |
                Select-Object -First 1

            if ($tab -and $tab.webSocketDebuggerUrl) {
                return $tab
            }
        } catch {}

        Start-Sleep -Milliseconds $RetryDelayMs
    }

    throw "No se encontró una pestaña target dentro del tiempo esperado."
}

try {
    Write-Host "--- Script Start: $(Get-Date) ---"

    if (-not (Test-Path $tempProfile)) {
        New-Item -ItemType Directory -Path $tempProfile | Out-Null
    }

    $launchTime = Get-Date
    $edgeArgs = "--remote-debugging-port=$global:DevToolsPort --user-data-dir=`"$tempProfile`" --no-first-run --new-window `"$TargetURL`""
    $edgeProcess = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru

    Start-Sleep -Seconds 5

    $tab = $null
    $socket = $null

    do {
        Start-Sleep -Milliseconds 700
        try {
            $tabs = Invoke-RestMethod -Uri "http://localhost:$global:DevToolsPort/json" -ErrorAction Stop

            Write-Host "`n--- DevTools targets ---"
            $tabs | ForEach-Object {
                Write-Host ("id={0} url={1} ws={2}" -f $_.id, $_.url, ($_.webSocketDebuggerUrl -as [string]))
            }
            Write-Host "------------------------`n"

            $tab = $tabs | Where-Object {
                $_.url -and (
                    $_.url -like "$TargetURL*" -or
                    $_.url -like "$PortalURL*"
                )
            } | Select-Object -First 1
        } catch {}
    } while (-not $tab)

    if ($tab.url -like "$PortalURL*") {
        Write-Host "Portal detectado. Esperando que el usuario abra la opción que lleva a la pestaña target..."
        $tab = Wait-ForTargetTab -TargetUrl $TargetURL -Port $global:DevToolsPort -MaxWaitSec 300
    }

    $socket = New-Object System.Net.WebSockets.ClientWebSocket
    $socket.ConnectAsync([Uri]$tab.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait(5000)

    $currentUrl = Wait-TabReady -Socket $socket -ExpectedUrl $TargetURL
    Write-Host "Pestaña lista: $currentUrl"

    Close-OtherTabs -KeepTabId $tab.id -Port $global:DevToolsPort

    Write-Host "Esperando validación de sesión..."

    do {
        $cookieResp = Send-CDPCommand $socket @{
            method = "Network.getCookies"
            params = @{}
        }

        $cookies = $cookieResp.result.cookies
        $found = $cookies | Where-Object { $_.name -eq "JSESSIONID" }

        if (-not $found) {
            Start-Sleep -Milliseconds 1500
        }
    } while (-not $found)

    Write-Host "Sesión lista."

    $cookieString = (
        $cookies | ForEach-Object { "$($_.name)=$($_.value)" }
    ) -join ";"

    $finalExpiry = (Get-Date).AddHours(1)
    $utf8NoBOM = New-Object System.Text.UTF8Encoding($false)

    [System.IO.File]::WriteAllText($SavePath, $cookieString, $utf8NoBOM)
    [System.IO.File]::WriteAllText($EXPIRY, $finalExpiry.ToString("MM/dd/yyyy HH:mm:ss"), $utf8NoBOM)

    Write-Host "Cookies guardadas."
}
catch {
    Write-Host "Error: $_"
    $CloseEdge = $true
}
finally {
    if ($socket -is [System.Net.WebSockets.ClientWebSocket] -and $socket.State -eq 'Open') {
        try {
            $socket.CloseAsync(
                [System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,
                "",
                [Threading.CancellationToken]::None
            ).Wait(3000)
        } catch {}
    }

    if ($CloseEdge) {
        try {
            Get-CimInstance Win32_Process |
                Where-Object { $_.CommandLine -and $_.CommandLine -like "*--remote-debugging-port=$global:DevToolsPort*" } |
                ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
        } catch {}
    }

    Stop-Transcript | Out-Null
    Start-Sleep -Seconds 2
    Debug-Hide-EdgeWindow
}