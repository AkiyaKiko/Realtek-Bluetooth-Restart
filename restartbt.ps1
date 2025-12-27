<# 
BT-Fix.ps1
- Restart Bluetooth Support Service (bthserv)
- Disable/Enable Bluetooth adapter (Realtek by default)
- Optionally trigger device rescan via pnputil
- Retries + logging
#>

[CmdletBinding()]
param(
    # 用正则匹配你的蓝牙适配器名称（默认匹配 Realtek）
    [string]$AdapterNameRegex = "Realtek.*Bluetooth",

    # 重试次数（有时第一次唤醒后禁用/启用会失败，重试更稳）
    [int]$Retries = 2,

    # 每次 Disable/Enable 之间等待秒数
    [int]$ToggleDelaySeconds = 2,

    # 重启 bthserv（蓝牙支持服务）
    [switch]$RestartBluetoothService = $true,

    # 是否触发一次“扫描硬件改动”（需要 pnputil，Win10/11 通常自带）
    [switch]$RescanDevices = $true,

    # 可选：写日志到文件
    [string]$LogPath = ""
)

function Write-Log {
    param([string]$Msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Msg"
    Write-Host $line
    if ($LogPath) {
        try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch {}
    }
}

function Assert-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "请用【管理员】运行 PowerShell（右键 → 以管理员身份运行），否则无法禁用/启用设备。"
    }
}

function Restart-BthServ {
    try {
        $svc = Get-Service -Name "bthserv" -ErrorAction Stop
        if ($svc.Status -eq "Running") {
            Write-Log "Restarting service: bthserv"
            Restart-Service -Name "bthserv" -Force -ErrorAction Stop
        } else {
            Write-Log "Starting service: bthserv (current: $($svc.Status))"
            Start-Service -Name "bthserv" -ErrorAction Stop
        }
        Start-Sleep -Seconds 1
    } catch {
        Write-Log "WARN: Failed to restart/start bthserv: $($_.Exception.Message)"
    }
}

function Find-BluetoothAdapter {
    param([string]$NameRegex)

    # 在 Bluetooth 类里找匹配的适配器
    $candidates = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
        Where-Object { $_.FriendlyName -match $NameRegex }

    if (-not $candidates) {
        # 兜底：有些驱动把名字写得怪，尝试更宽松的查找
        $candidates = Get-PnpDevice -ErrorAction SilentlyContinue |
            Where-Object { $_.FriendlyName -match $NameRegex -and $_.Class -match "Bluetooth" }
    }

    # 优先选择 “Adapter” / “Radio” 这种看起来像主设备的条目
    $best = $candidates | Sort-Object @{
        Expression = {
            $n = $_.FriendlyName
            if ($n -match "Adapter|Radio") { 0 } else { 1 }
        }
    } | Select-Object -First 1

    return $best
}

function Toggle-PnpDevice {
    param(
        [Parameter(Mandatory=$true)] $Device,
        [int]$DelaySeconds = 2
    )

    Write-Log "Disabling: $($Device.FriendlyName)"
    Disable-PnpDevice -InstanceId $Device.InstanceId -Confirm:$false -ErrorAction Stop
    Start-Sleep -Seconds $DelaySeconds

    Write-Log "Enabling:  $($Device.FriendlyName)"
    Enable-PnpDevice -InstanceId $Device.InstanceId -Confirm:$false -ErrorAction Stop
    Start-Sleep -Seconds 1
}

function Rescan-Devices {
    # pnputil /scan-devices 在 Win10/11 一般可用，用来触发一次重新枚举
    try {
        $pnputil = (Get-Command pnputil.exe -ErrorAction Stop).Source
        Write-Log "Triggering device rescan: pnputil /scan-devices"
        & $pnputil /scan-devices | Out-Null
        Start-Sleep -Seconds 2
    } catch {
        Write-Log "WARN: pnputil not available or scan failed: $($_.Exception.Message)"
    }
}

try {
    Assert-Admin
    if ($LogPath) {
        Write-Log "Logging to: $LogPath"
    }

    Write-Log "=== Bluetooth Fix Start ==="
    Write-Log "AdapterNameRegex = $AdapterNameRegex, Retries = $Retries"

    if ($RestartBluetoothService) {
        Restart-BthServ
    }

    $adapter = Find-BluetoothAdapter -NameRegex $AdapterNameRegex
    if (-not $adapter) {
        Write-Log "ERROR: 没找到匹配的蓝牙适配器（regex: $AdapterNameRegex）"
        Write-Log "提示：你可以先运行：Get-PnpDevice -Class Bluetooth | Select FriendlyName,Status"
        exit 1
    }

    Write-Log "Target adapter: $($adapter.FriendlyName)  Status=$($adapter.Status)"

    $attempt = 0
    $ok = $false
    while ($attempt -le $Retries -and -not $ok) {
        $attempt++
        try {
            Write-Log "--- Attempt $attempt/$Retries ---"
            Toggle-PnpDevice -Device $adapter -DelaySeconds $ToggleDelaySeconds

            if ($RescanDevices) {
                Rescan-Devices
            }

            # 再读一次状态
            $adapter2 = Get-PnpDevice -InstanceId $adapter.InstanceId -ErrorAction SilentlyContinue
            if ($adapter2 -and $adapter2.Status -eq "OK") {
                Write-Log "SUCCESS: Adapter is OK."
                $ok = $true
            } else {
                Write-Log "WARN: Adapter status is '$($adapter2.Status)'."
            }
        } catch {
            Write-Log "WARN: Attempt failed: $($_.Exception.Message)"
            Start-Sleep -Seconds 2
        }
    }

    if (-not $ok) {
        Write-Log "FINAL: 仍未确认恢复（但有时即使状态没变，耳机也能重新连上）。建议此时尝试重新连接耳机。"
        exit 2
    }

    Write-Log "=== Bluetooth Fix Done ==="
    exit 0
}
catch {
    Write-Host $_.Exception.Message
    exit 99
}
