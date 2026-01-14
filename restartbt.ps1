[CmdletBinding()]
param(
    [string]$LogPath = "",
    [int]$PnPTimeoutSec = 20,
    [int]$SvcTimeoutSec = 15
)

function Write-Log {
    param([string]$Msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Msg"
    Write-Host $line
    if ($LogPath) { try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch {} }
}

function Assert-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "请用【管理员】运行 PowerShell（右键 → 以管理员身份运行）。"
    }
}

function Wait-WithTimeout {
    param(
        [scriptblock]$Action,
        [int]$TimeoutSec,
        [string]$What = "operation"
    )
    $job = Start-Job -ScriptBlock $Action
    try {
        if (-not (Wait-Job -Job $job -Timeout $TimeoutSec)) {
            Stop-Job $job -Force | Out-Null
            throw "$What timed out after $TimeoutSec sec."
        }
        $out = Receive-Job $job -ErrorAction SilentlyContinue
        return $out
    } finally {
        Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

function Restart-BluetoothPnP {
    param([int]$TimeoutSec)

    # 更稳：先找 Bluetooth class 的设备，再优先挑 Realtek，其次挑 “Bluetooth Adapter/Radio”
    $candidates = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue
    if (-not $candidates) {
        Write-Log "No PnP devices found in Class=Bluetooth."
        return $false
    }

    $dev =
        $candidates | Where-Object { $_.FriendlyName -match "Realtek" } | Select-Object -First 1
    if (-not $dev) {
        $dev = $candidates | Where-Object { $_.FriendlyName -match "Adapter|Radio|Bluetooth" } | Select-Object -First 1
    }
    if (-not $dev) {
        $dev = $candidates | Select-Object -First 1
    }

    Write-Log "PnP target: $($dev.FriendlyName) | Status=$($dev.Status) | InstanceId=$($dev.InstanceId)"

    try {
        Write-Log "Disabling PnP device..."
        Wait-WithTimeout -TimeoutSec $TimeoutSec -What "Disable-PnpDevice" -Action {
            param($id)
            Disable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        }.GetNewClosure() | Out-Null
    } catch {
        # 有些机器 Disable 会报正在使用/拒绝，但 Enable 仍可能恢复，所以不中断
        Write-Log "WARN: Disable may have failed: $($_.Exception.Message)"
    }

    Start-Sleep -Seconds 2

    try {
        Write-Log "Enabling PnP device..."
        Wait-WithTimeout -TimeoutSec $TimeoutSec -What "Enable-PnpDevice" -Action {
            param($id)
            Enable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        }.GetNewClosure() | Out-Null
    } catch {
        Write-Log "ERROR: Enable failed: $($_.Exception.Message)"
        return $false
    }

    Start-Sleep -Seconds 1
    $dev2 = Get-PnpDevice -InstanceId $dev.InstanceId -ErrorAction SilentlyContinue
    Write-Log "PnP after: Status=$($dev2.Status)"
    return $true
}

function Restart-BthServ {
    param([int]$TimeoutSec)

    try {
        $svc = Get-Service -Name "bthserv" -ErrorAction Stop
        Write-Log "bthserv current: $($svc.Status)"

        if ($svc.Status -eq "Running") {
            Write-Log "Restarting service: bthserv"
            Wait-WithTimeout -TimeoutSec $TimeoutSec -What "Restart-Service bthserv" -Action {
                Restart-Service -Name "bthserv" -Force -ErrorAction Stop
            } | Out-Null
        } else {
            Write-Log "Starting service: bthserv"
            Wait-WithTimeout -TimeoutSec $TimeoutSec -What "Start-Service bthserv" -Action {
                Start-Service -Name "bthserv" -ErrorAction Stop
            } | Out-Null
        }

        Start-Sleep -Seconds 1
        $svc2 = Get-Service -Name "bthserv" -ErrorAction SilentlyContinue
        Write-Log "bthserv after: $($svc2.Status)"
        return $true
    } catch {
        Write-Log "ERROR: bthserv op failed: $($_.Exception.Message)"
        return $false
    }
}

try {
    Assert-Admin
    if ($LogPath) { Write-Log "Logging to: $LogPath" }

    Write-Log "=== Bluetooth Fix Start (PnP + Service) ==="

    $pnpOk = Restart-BluetoothPnP -TimeoutSec $PnPTimeoutSec
    if (-not $pnpOk) {
        Write-Log "PnP restart failed; still trying service restart..."
    }

    $svcOk = Restart-BthServ -TimeoutSec $SvcTimeoutSec

    if ($pnpOk -or $svcOk) {
        Write-Log "=== Bluetooth Fix Done ==="
        exit 0
    } else {
        Write-Log "=== Bluetooth Fix Failed ==="
        exit 2
    }
}
catch {
    Write-Host $_.Exception.Message
    exit 99
}
