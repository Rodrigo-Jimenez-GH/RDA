param(
    [string]$SavePath,
    [string]$EXPIRY,
    [bool]$CloseEdge = $false
)

#param(
#    [string]$SavePath = "C:\Macros\RESPS\TOKEN.txt",
#    [string]$EXPIRY = "C:\Macros\RESPS\TOKEN_EOPS.txt",
#    [bool]$CloseEdge = $false
#)


$logPath = "$env:TEMP\FGC_Cookie.log"
Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue

$TargetURL = "https://fgc-lac-cairo-atl.prod.cloud.fedex.com/clearance/mainMenu.jsp"
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

function Log {
    param([string]$Msg)
    $ts = Get-Date -Format "HH:mm:ss.fff"
    Write-Host "[$ts] $Msg"
}

function Debug-Hide-EdgeWindow {
    param(
        [int]$Pid
    )

    try {

        $p = Get-Process -Id $Pid -ErrorAction Stop

        if ($p.MainWindowHandle -ne 0) {

            $null = [WinAPI]::ShowWindowAsync(
                $p.MainWindowHandle,
                0
            )

            Log "Hidden PID=$Pid"
        }
        else {
            Log "PID=$Pid has no visible window"
        }

    }
    catch {
        Log "Hide failed: $_"
    }
}

function Send-CDPCommand {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
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

    $Socket.SendAsync(
        [ArraySegment[byte]]$bytes,
        [System.Net.WebSockets.WebSocketMessageType]::Text,
        $true,
        [Threading.CancellationToken]::None
    ).Wait()

    $buffer = New-Object byte[] 102400

    $result = $Socket.ReceiveAsync(
        [ArraySegment[byte]]$buffer,
        [Threading.CancellationToken]::None
    ).Result

    $text = [Text.Encoding]::UTF8.GetString($buffer,0,$result.Count)

    $obj = $text | ConvertFrom-Json

    if ($obj -is [System.Array]) {
        return $obj[-1]
    }

    return $obj
}

function Wait-TabReady {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [string]$ExpectedUrl,
        [int]$RetryDelayMs = 500,
        [int]$StableHitsRequired = 2
    )

    $stableHits = 0

    while ($true) {

        $urlResp = Send-CDPCommand $Socket @{
            method = "Runtime.evaluate"
            params = @{ expression = "window.location.href" }
        }

        $stateResp = Send-CDPCommand $Socket @{
            method = "Runtime.evaluate"
            params = @{ expression = "document.readyState" }
        }

        $cookieResp = Send-CDPCommand $Socket @{
            method = "Network.getCookies"
            params = @{}
        }

        $currentUrl = $urlResp.result.result.value
        $state = $stateResp.result.result.value
        $cookies = $cookieResp.result.cookies
        $hasSession = $cookies.name -contains "JSESSIONID"

        Log "URL=$currentUrl"
        Log "STATE=$state"
        Log "SESSION=$hasSession"

        $good =
            ($currentUrl -like "$ExpectedUrl*") -and
            ($state -eq "complete") -and
            $hasSession

        if ($good) {
            $stableHits++
            Log "Stable hit $stableHits/$StableHitsRequired"
        }
        else {
            $stableHits = 0
        }

        if ($stableHits -ge $StableHitsRequired) {
            Log "Tab stable"
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

    Log "Close-OtherTabs START"

    $tabs = Invoke-RestMethod "http://localhost:$Port/json"

    Log "Found $($tabs.Count) tabs"

    foreach ($t in $tabs) {

        Log "Inspect $($t.id) -> $($t.url)"

if (-not $t.id) {
    Log "Skip no-id"
    continue
}

if (-not $t.url) {
    Log "Skip empty-url target"
    continue
}

if ($t.url -like "devtools://*") {
    Log "Skip devtools"
    continue
}

if ($t.type -ne "page") {
    Log "Skip non-page target ($($t.type))"
    continue
}

        if ($t.id -eq $KeepTabId) {
            Log "Keep target"
            continue
        }

        try {
            Invoke-RestMethod `
                "http://localhost:$Port/json/close/$($t.id)" `
                -Method GET `
                -TimeoutSec 5 | Out-Null

            Log "Closed"
        }
        catch {
            Log "Close failed: $_"
        }
    }

    Log "Close-OtherTabs END"
}

function Wait-ForTargetTab {
    param(
        [string]$TargetUrl,
        [int]$Port = $global:DevToolsPort,
        [int]$MaxWaitSec = 300
    )

    $sw = [Diagnostics.Stopwatch]::StartNew()

    while ($sw.Elapsed.TotalSeconds -lt $MaxWaitSec) {

        try {

            $tabs = Invoke-RestMethod "http://localhost:$Port/json"

            Log "Scanning tabs..."

            foreach ($t in $tabs) {
                Log "TAB $($t.id) $($t.url)"
            }

            $tab = $tabs |
                Where-Object {
                    $_.url -and $_.url -like "$TargetUrl*"
                } |
                Select-Object -First 1

            if ($tab) {
                Log "Target tab detected"
                return $tab
            }

        } catch {
            Log "Scan failed: $_"
        }

        Start-Sleep -Seconds 2
    }

    throw "No target tab detected"
}

try {

    Log "SCRIPT START"

    if (-not (Test-Path $tempProfile)) {
        New-Item -ItemType Directory -Path $tempProfile | Out-Null
    }

    $edgeArgs =
        "--remote-debugging-port=$global:DevToolsPort " +
        "--user-data-dir=`"$tempProfile`" " +
        "--no-first-run --new-window `"$TargetURL`""

    Log "Launching Edge"

    $edgeProcess = Start-Process `
        -FilePath $edge `
        -ArgumentList $edgeArgs `
        -PassThru

    Log "PID=$($edgeProcess.Id)"

    Start-Sleep -Milliseconds 800

    do {

        Start-Sleep -Milliseconds 700

        $tabs = Invoke-RestMethod `
            "http://localhost:$global:DevToolsPort/json"

        foreach ($t in $tabs) {
            Log "TAB $($t.id) $($t.url)"
        }

        $tab = $tabs |
            Where-Object {
                $_.url -like "$TargetURL*" -or
                $_.url -like "$PortalURL*"
            } |
            Select-Object -First 1

    } while (-not $tab)

    if ($tab.url -like "$PortalURL*") {

        Log "Portal detected. Waiting target..."

        $tab = Wait-ForTargetTab `
            -TargetUrl $TargetURL
    }

    Log "Connecting socket"

    $socket = New-Object System.Net.WebSockets.ClientWebSocket

    $null = $socket.ConnectAsync(
        [Uri]$tab.webSocketDebuggerUrl,
        [Threading.CancellationToken]::None
    ).Wait(5000)

    Log "Socket connected"

    $currentUrl = Wait-TabReady `
        -Socket $socket `
        -ExpectedUrl $TargetURL

    Log "READY: $currentUrl"

    Close-OtherTabs -KeepTabId $tab.id

    Log "Extract cookies"

    $cookieResp = Send-CDPCommand $socket @{
        method = "Network.getCookies"
        params = @{}
    }

    $cookies = $cookieResp.result.cookies

    $cookieString = (
        $cookies |
        ForEach-Object { "$($_.name)=$($_.value)" }
    ) -join ";"

    $finalExpiry = (Get-Date).AddHours(1)

    $utf8NoBOM =
        New-Object System.Text.UTF8Encoding($false)

    [System.IO.File]::WriteAllText(
        $SavePath,
        $cookieString,
        $utf8NoBOM
    )

    [System.IO.File]::WriteAllText(
        $EXPIRY,
        $finalExpiry.ToString("MM/dd/yyyy HH:mm:ss"),
        $utf8NoBOM
    )

    Log "Cookies saved"
}
catch {

    Log "ERROR: $_"
    $CloseEdge = $true
}
finally {

    if (
        $socket -and
        $socket.State -eq 'Open'
    ) {
        try {
            Log "Closing socket"

            $null = $socket.CloseAsync(
                [System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,
                "",
                [Threading.CancellationToken]::None
            ).Wait(3000)
        }
        catch {
            Log "Socket close failed"
        }
    }

    if ($CloseEdge -and $edgeProcess) {
        try {
            Log "Killing Edge PID=$($edgeProcess.Id)"
            Stop-Process -Id $edgeProcess.Id -Force
            Log "Edge killed"
        }
        catch {
            Log "Kill failed: $_"
        }
    }

    Stop-Transcript | Out-Null

    Debug-Hide-EdgeWindow
}