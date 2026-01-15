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
    if ($LogPath) {
        try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch {}
    }
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
        [object[]]$ArgumentList = @(),
        [int]$TimeoutSec,
        [string]$What = "operation"
    )

    # 关键修复：把参数传给 Job，否则 param($id) 会是 $null
    $job = Start-Job -ScriptBlock $Action -ArgumentList $ArgumentList

    try {
        if (-not (Wait-Job -Job $job -Timeout $TimeoutSec)) {
            Stop-Job $job -Force | Out-Null
            throw "$What timed out after $TimeoutSec sec."
        }

        # 把 Job 中的输出/错误都收回来（错误会进入 error stream，但这里避免中断）
        $out = Receive-Job $job -ErrorAction SilentlyContinue
        return $out
    }
    finally {
        Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

function Get-BluetoothTargetDevice {
    # 更稳的选法：
    # 1) 优先 Realtek + (Radio/Adapter)
    # 2) 再 Realtek 任意
    # 3) 再 Radio/Adapter
    # 4) 最后随便挑一个 Bluetooth class 里的
    $candidates = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue
    if (-not $candidates) { return $null }

    $dev = $candidates | Where-Object { $_.FriendlyName -match "Realtek" -and $_.FriendlyName -match "Radio|Adapter" } | Select-Object -First 1
    if (-not $dev) { $dev = $candidates | Where-Object { $_.FriendlyName -match "Realtek" } | Select-Object -First 1 }
    if (-not $dev) { $dev = $candidates | Where-Object { $_.FriendlyName -match "Radio|Adapter|Bluetooth" } | Select-Object -First 1 }
    if (-not $dev) { $dev = $candidates | Select-Object -First 1 }

    return $dev
}

function Restart-BluetoothPnP {
    param([int]$TimeoutSec)

    $dev = Get-BluetoothTargetDevice
    if (-not $dev) {
        Write-Log "No PnP devices found in Class=Bluetooth."
        return $false
    }

    Write-Log "PnP target: $($dev.FriendlyName) | Status=$($dev.Status) | InstanceId=$($dev.InstanceId)"

    # 打印重启前状态
    $before = Get-PnpDevice -InstanceId $dev.InstanceId -ErrorAction SilentlyContinue
    if ($before) {
        Write-Log ("PnP before: Status={0} ProblemCode={1}" -f $before.Status, $before.ProblemCode)
    }

    try {
        Write-Log "Disabling PnP device..."
        Wait-WithTimeout -TimeoutSec $TimeoutSec -What "Disable-PnpDevice" -ArgumentList @($dev.InstanceId) -Action {
            param($id)
            Disable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        } | Out-Null
    } catch {
        # 有些机器 Disable 会报正在使用/拒绝，但 Enable 仍可能恢复，所以不中断
        Write-Log "WARN: Disable may have failed: $($_.Exception.Message)"
    }

    Start-Sleep -Seconds 2

    try {
        Write-Log "Enabling PnP device..."
        Wait-WithTimeout -TimeoutSec $TimeoutSec -What "Enable-PnpDevice" -ArgumentList @($dev.InstanceId) -Action {
            param($id)
            Enable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        } | Out-Null
    } catch {
        Write-Log "ERROR: Enable failed: $($_.Exception.Message)"
        return $false
    }

    Start-Sleep -Seconds 1

    $after = Get-PnpDevice -InstanceId $dev.InstanceId -ErrorAction SilentlyContinue
    if ($after) {
        Write-Log ("PnP after: Status={0} ProblemCode={1}" -f $after.Status, $after.ProblemCode)
        # 一般 ProblemCode=0 且 Status=OK 比较正常
        return $true
    } else {
        Write-Log "WARN: Device not found after enable (may have re-enumerated)."
        return $true
    }
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
