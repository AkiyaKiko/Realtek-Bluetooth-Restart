#Requires -RunAsAdministrator
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$TargetFriendlyName    = "Realtek Bluetooth Adapter"
$MaxRetries            = 5
$RetryDelaySeconds     = 2
$StateTimeoutSeconds   = 15
$AllowFuzzyMatch       = $true

function Get-DeviceByNameOrFuzzy {
    param([string]$FriendlyName, [bool]$Fuzzy)

    $dev = Get-PnpDevice -FriendlyName $FriendlyName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($dev) { return $dev }

    if (-not $Fuzzy) { return $null }

    # 优先 Bluetooth 类
    $dev = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
        Where-Object { $_.FriendlyName -like "*$FriendlyName*" } |
        Select-Object -First 1
    if ($dev) { return $dev }

    # 最后兜底：全局模糊
    return (Get-PnpDevice -ErrorAction SilentlyContinue |
        Where-Object { $_.FriendlyName -like "*$FriendlyName*" } |
        Select-Object -First 1)
}

function Get-InstancePrefix {
    param([string]$InstanceId)
    if ($InstanceId -match '^(.*\\VID_[0-9A-Fa-f]{4}&PID_[0-9A-Fa-f]{4})\\') { return ($Matches[1] + "\") }
    $idx = $InstanceId.LastIndexOf("\")
    if ($idx -gt 0) { return $InstanceId.Substring(0, $idx + 1) }
    return $InstanceId
}

function Find-DeviceResilient {
    param(
        [string]$PreferredInstanceId,
        [string]$InstancePrefix,
        [string]$FriendlyName
    )

    $dev = Get-PnpDevice -InstanceId $PreferredInstanceId -ErrorAction SilentlyContinue
    if ($dev) { return $dev }
    if ($InstancePrefix) {
        $dev = Get-PnpDevice -ErrorAction SilentlyContinue |
            Where-Object { $_.InstanceId -like "$InstancePrefix*" } |
            Sort-Object -Property Status -Descending |
            Select-Object -First 1
        if ($dev) { return $dev }
    }

    return (Get-DeviceByNameOrFuzzy -FriendlyName $FriendlyName -Fuzzy:$AllowFuzzyMatch)
}

function Wait-ForStatus {
    param(
        [string]$InstanceId,
        [string]$ExpectedStatus,
        [int]$TimeoutSeconds = 15
    )

    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $cur = Get-PnpDevice -InstanceId $InstanceId -ErrorAction SilentlyContinue
        if ($null -eq $cur) {
            Start-Sleep -Milliseconds 250
            continue
        }
        if ($cur.Status -eq $ExpectedStatus) { return $true }
        Start-Sleep -Milliseconds 250
    }
    return $false
}

function Test-DisableSmart {
    param([string]$InstanceId, [int]$TimeoutSeconds)

    $cur = Get-PnpDevice -InstanceId $InstanceId -ErrorAction SilentlyContinue
    if ($cur -and $cur.Status -eq "Disabled") { return $true }

    $threw = $false
    $throwMsg = $null

    try {
        Disable-PnpDevice -InstanceId $InstanceId -Confirm:$false -ErrorAction Stop
    } catch {
        $threw = $true
        $throwMsg = $_.Exception.Message

        Write-Host "Disable reported: $throwMsg (will verify by status)" -ForegroundColor DarkGray
    }

    if (Wait-ForStatus -InstanceId $InstanceId -ExpectedStatus "Disabled" -TimeoutSeconds $TimeoutSeconds) {
        return $true
    }

    $cur2 = Get-PnpDevice -InstanceId $InstanceId -ErrorAction SilentlyContinue
    if ($cur2 -and $cur2.Status -eq "Disabled") { return $true }

    if ($threw -and $cur2 -and $cur2.Status -eq "OK") {
        return $false  # 交给上层走 “Enable anyway / 软恢复” 路径，但不必再额外 warning
    }

    if ($threw) {
        Write-Warning "Disable failed and device is not Disabled/OK. Last status: $($cur2.Status)"
    }

    return $false
}


function Test-EnableSmart {
    param([string]$InstanceId, [int]$TimeoutSeconds)

    $cur = Get-PnpDevice -InstanceId $InstanceId -ErrorAction SilentlyContinue
    if ($cur -and $cur.Status -eq "OK") { return $true }

    $threw = $false
    $throwMsg = $null

    try {
        Enable-PnpDevice -InstanceId $InstanceId -Confirm:$false -ErrorAction Stop
    } catch {
        $threw = $true
        $throwMsg = $_.Exception.Message
        Write-Host "Enable reported: $throwMsg (will verify by status)" -ForegroundColor DarkGray
    }

    if (Wait-ForStatus -InstanceId $InstanceId -ExpectedStatus "OK" -TimeoutSeconds $TimeoutSeconds) {
        return $true
    }

    $cur2 = Get-PnpDevice -InstanceId $InstanceId -ErrorAction SilentlyContinue
    if ($cur2 -and $cur2.Status -eq "OK") { return $true }

    if ($threw) {
        Write-Warning "Enable failed and device did not reach OK. Last status: $($cur2.Status)"
    }

    return $false
}


$first = Get-DeviceByNameOrFuzzy -FriendlyName $TargetFriendlyName -Fuzzy:$AllowFuzzyMatch
if (-not $first) {
    Write-Host "Bluetooth device not found: $TargetFriendlyName" -ForegroundColor Red
    exit 1
}

$preferredInstanceId = $first.InstanceId
$instancePrefix = Get-InstancePrefix -InstanceId $preferredInstanceId

Write-Host "Found device: $($first.FriendlyName)" -ForegroundColor Cyan
Write-Host "InstanceId : $preferredInstanceId" -ForegroundColor Cyan
Write-Host "Status     : $($first.Status)" -ForegroundColor Cyan

for ($i = 1; $i -le $MaxRetries; $i++) {
    $dev = Find-DeviceResilient -PreferredInstanceId $preferredInstanceId -InstancePrefix $instancePrefix -FriendlyName $TargetFriendlyName
    if (-not $dev) {
        Write-Warning "[$i/$MaxRetries] Device temporarily not found; waiting..."
        Start-Sleep -Seconds $RetryDelaySeconds
        continue
    }

    $id = $dev.InstanceId
    $status = $dev.Status
    Write-Host "[$i/$MaxRetries] Current status: $status" -ForegroundColor DarkGray

    if ($status -eq "Disabled") {
        Write-Host "[$i/$MaxRetries] Device is already Disabled -> Enabling..." -ForegroundColor Yellow
        if (Test-EnableSmart -InstanceId $id -TimeoutSeconds $StateTimeoutSeconds) {
            Write-Host "Bluetooth driver restart complete (Disabled->Enabled)." -ForegroundColor Green
            exit 0
        }
        Write-Warning "[$i/$MaxRetries] Enable did not reach OK; retrying..."
        Start-Sleep -Seconds $RetryDelaySeconds
        continue
    }

    if ($status -eq "Error") {
        Write-Host "[$i/$MaxRetries] Status=Error -> Trying soft recovery (Enable only)..." -ForegroundColor Yellow
        if (Test-EnableSmart -InstanceId $id -TimeoutSeconds $StateTimeoutSeconds) {
            Write-Host "Soft recovery succeeded (OK after Enable). Stop to avoid breaking a working connection." -ForegroundColor Green
            exit 0
        }
        Write-Warning "[$i/$MaxRetries] Soft recovery failed; falling back to Disable->Enable..."
    }

    Write-Host "[$i/$MaxRetries] Disabling device..." -ForegroundColor Yellow
    $disabled = Test-DisableSmart -InstanceId $id -TimeoutSeconds $StateTimeoutSeconds

    if (-not $disabled) {
        Write-Warning "[$i/$MaxRetries] Disable did not reach Disabled. Trying Enable anyway to avoid making it worse..."
        if (Test-EnableSmart -InstanceId $id -TimeoutSeconds $StateTimeoutSeconds) {
            Write-Host "Bluetooth seems healthy (OK after Enable). Stop here to avoid breaking a working connection." -ForegroundColor Green
            exit 0
        }

        Write-Warning "[$i/$MaxRetries] Still not OK; retrying..."
        Start-Sleep -Seconds $RetryDelaySeconds
        continue
    }

    Write-Host "[$i/$MaxRetries] Enabling device..." -ForegroundColor Yellow
    if (Test-EnableSmart -InstanceId $id -TimeoutSeconds $StateTimeoutSeconds) {
        Write-Host "Bluetooth driver restart complete." -ForegroundColor Green
        exit 0
    }

    Write-Warning "[$i/$MaxRetries] Enable did not reach OK; retrying..."
    Start-Sleep -Seconds $RetryDelaySeconds
}

throw "Failed to restart device after $MaxRetries attempts."

