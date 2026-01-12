[CmdletBinding()]
param(
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
        throw "请用【管理员】运行 PowerShell（右键 → 以管理员身份运行），否则可能无法操作服务。"
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
        $svc2 = Get-Service -Name "bthserv" -ErrorAction SilentlyContinue
        Write-Log "bthserv status: $($svc2.Status)"
    } catch {
        Write-Log "ERROR: Failed to restart/start bthserv: $($_.Exception.Message)"
        exit 2
    }
}

try {
    Assert-Admin
    if ($LogPath) { Write-Log "Logging to: $LogPath" }

    Write-Log "=== Bluetooth Fix Start (Service-only) ==="
    Restart-BthServ
    Write-Log "=== Bluetooth Fix Done ==="
    exit 0
}
catch {
    Write-Host $_.Exception.Message
    exit 99
}
