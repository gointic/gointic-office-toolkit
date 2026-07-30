#requires -Version 5.1
#requires -RunAsAdministrator
<#
.SYNOPSIS
  GOINTIC Office Repair Tool v2.0

.DESCRIPTION
  Herramienta técnica independiente para diagnosticar, reparar y limpiar
  instalaciones de Microsoft Office en Windows 10/11.

  IMPORTANTE:
  - No es OffScrub ni una herramienta oficial de Microsoft.
  - No elimina documentos, PST/OST, OneDrive, Edge ni Visual C++.
  - Puede eliminar Office, Click-to-Run, tareas, servicios y claves específicas.
  - Crea logs, copias de registro e informe HTML.

.OPCIONES
  1. Diagnóstico de Office
  2. Reparación rápida de Click-to-Run
  3. Limpieza completa de Office
  4. Ejecutar DISM y SFC
  5. Reinstalar Office mediante ODT existente
  6. Generar informe HTML
  7. Salir

.PARAMETER Silent
  Ejecuta diagnóstico y genera informe sin mostrar menú.

.PARAMETER Force
  Omite confirmaciones en la limpieza.

.PARAMETER PurgeUserSettings
  Elimina configuración y cachés de Office del usuario actual.
  No elimina PST, OST, documentos ni plantillas personales.

.PARAMETER RemoveTeamsAddin
  Elimina también Microsoft Teams Meeting Add-in for Microsoft Office.

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File .\GOINTIC-Office-Repair-Tool-v2.ps1

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File .\GOINTIC-Office-Repair-Tool-v2.ps1 -Silent
#>

[CmdletBinding()]
param(
    [switch]$Silent,
    [switch]$Force,
    [switch]$PurgeUserSettings,
    [switch]$RemoveTeamsAddin
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$ToolVersion = '2.0'
$Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$BaseDir = Join-Path $env:SystemDrive "GOINTIC-OfficeRepair-$Timestamp"
$LogFile = Join-Path $BaseDir 'GOINTIC-OfficeRepair.log'
$ReportFile = Join-Path $BaseDir 'GOINTIC-OfficeRepair-Report.html'
$RegistryBackupDir = Join-Path $BaseDir 'Registro'
$script:DiagnosticData = [ordered]@{}

New-Item -Path $BaseDir, $RegistryBackupDir -ItemType Directory -Force | Out-Null
Start-Transcript -Path $LogFile -Force | Out-Null

function Write-Title {
    Clear-Host
    Write-Host '=====================================================' -ForegroundColor DarkCyan
    Write-Host " GOINTIC Office Repair Tool v$ToolVersion" -ForegroundColor Cyan
    Write-Host '=====================================================' -ForegroundColor DarkCyan
    Write-Host " Equipo: $env:COMPUTERNAME"
    Write-Host " Log:    $LogFile"
    Write-Host ''
}

function Write-Step {
    param([Parameter(Mandatory)][string]$Text)
    Write-Host "`n=== $Text ===" -ForegroundColor Cyan
}

function Wait-Key {
    if (-not $Silent) {
        Write-Host ''
        Read-Host 'Pulsa ENTER para continuar' | Out-Null
    }
}

function Test-OfficeName {
    param([string]$DisplayName)
    if (-not $DisplayName) { return $false }

    $officeMatch = $DisplayName -match '(?i)\b(Microsoft\s+365|Microsoft\s+Office|Office\s+(Professional|Standard|Home|Hogar|LTSC)|Microsoft\s+(Visio|Project)|Office\s+16\s+Click-to-Run)'
    $excluded = $DisplayName -match '(?i)(OneDrive|Edge|Visual C\+\+|Update Health|WebView2|Teams Machine-Wide Installer)'
    if ($RemoveTeamsAddin) {
        $excluded = $excluded -and $DisplayName -notmatch '(?i)Teams Meeting Add-in'
    } else {
        $excluded = $excluded -or $DisplayName -match '(?i)Teams Meeting Add-in'
    }
    return ($officeMatch -and -not $excluded)
}

function Get-OfficeEntries {
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    Get-ItemProperty $roots -ErrorAction SilentlyContinue |
        Where-Object { Test-OfficeName $_.DisplayName } |
        Sort-Object DisplayName -Unique
}

function Export-RegKey {
    param(
        [Parameter(Mandatory)][string]$NativeKey,
        [Parameter(Mandatory)][string]$Name
    )
    $destination = Join-Path $RegistryBackupDir $Name
    & reg.exe export $NativeKey $destination /y 2>$null | Out-Null
}

function Remove-RegKeySafe {
    param([Parameter(Mandatory)][string]$Path)
    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write-Host "Clave eliminada: $Path"
        } catch {
            Write-Warning "No se pudo eliminar $Path | $($_.Exception.Message)"
        }
    }
}

function Remove-PathSafe {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }

    try {
        takeown.exe /F "$Path" /R /D Y 2>$null | Out-Null
        icacls.exe "$Path" /grant '*S-1-5-32-544:F' /T /C /Q 2>$null | Out-Null
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        Write-Host "Eliminado: $Path"
    } catch {
        Write-Warning "No se pudo eliminar $Path | $($_.Exception.Message)"
    }
}

function Test-BinaryHealth {
    param([Parameter(Mandatory)][string]$Path)

    $result = [ordered]@{
        Path = $Path
        Exists = $false
        Signature = 'No comprobada'
        Version = ''
        Healthy = $false
    }

    if (Test-Path -LiteralPath $Path) {
        $result.Exists = $true
        try {
            $signature = Get-AuthenticodeSignature -FilePath $Path
            $result.Signature = [string]$signature.Status
            $result.Version = (Get-Item $Path).VersionInfo.FileVersion
            $result.Healthy = ($signature.Status -eq 'Valid')
        } catch {
            $result.Signature = "Error: $($_.Exception.Message)"
        }
    }
    [pscustomobject]$result
}

function Get-OfficeDiagnostics {
    Write-Step 'Diagnóstico de Office'

    $entries = @(Get-OfficeEntries)
    $clickService = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
    $c2rFolder = "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun"
    $officeExe = Join-Path $c2rFolder 'OfficeClickToRun.exe'
    $c2rDll = Join-Path $c2rFolder 'C2RUI.dll'
    $winwordPaths = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
    ) | Where-Object { $_ -and (Test-Path $_) }

    $officeHealth = Test-BinaryHealth -Path $officeExe
    $dllHealth = Test-BinaryHealth -Path $c2rDll

    $script:DiagnosticData = [ordered]@{
        Fecha = (Get-Date)
        Equipo = $env:COMPUTERNAME
        Usuario = $env:USERNAME
        Windows = (Get-CimInstance Win32_OperatingSystem).Caption
        WindowsVersion = (Get-CimInstance Win32_OperatingSystem).Version
        OfficeEntries = $entries
        ClickToRunService = if ($clickService) { "$($clickService.Status) / $($clickService.StartType)" } else { 'No existe' }
        ClickToRunFolder = (Test-Path $c2rFolder)
        OfficeClickToRun = $officeHealth
        C2RUI = $dllHealth
        WinWordPaths = $winwordPaths
        RebootPending = (
            (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
            (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
        )
    }

    Write-Host "Windows: $($script:DiagnosticData.Windows) $($script:DiagnosticData.WindowsVersion)"
    Write-Host "Reinicio pendiente: $($script:DiagnosticData.RebootPending)"
    Write-Host "Servicio Click-to-Run: $($script:DiagnosticData.ClickToRunService)"
    Write-Host "Carpeta Click-to-Run: $($script:DiagnosticData.ClickToRunFolder)"
    Write-Host "OfficeClickToRun.exe: existe=$($officeHealth.Exists), firma=$($officeHealth.Signature)"
    Write-Host "C2RUI.dll: existe=$($dllHealth.Exists), firma=$($dllHealth.Signature)"

    if ($entries.Count -eq 0) {
        Write-Host 'No se detectan suites Office registradas.' -ForegroundColor Yellow
    } else {
        Write-Host 'Productos detectados:' -ForegroundColor Green
        $entries | ForEach-Object { Write-Host " - $($_.DisplayName) [$($_.DisplayVersion)]" }
    }

    if ($winwordPaths.Count -gt 0) {
        Write-Host 'WINWORD.EXE detectado en:'
        $winwordPaths | ForEach-Object { Write-Host " - $_" }
    }

    return $script:DiagnosticData
}

function Repair-ClickToRun {
    Write-Step 'Reparación rápida de Click-to-Run'

    $service = Get-Service ClickToRunSvc -ErrorAction SilentlyContinue
    $client = "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"

    if (-not $service -or -not (Test-Path $client)) {
        Write-Warning 'Click-to-Run no está instalado o está demasiado dañado para una reparación rápida.'
        return
    }

    try {
        Set-Service ClickToRunSvc -StartupType Automatic
        Start-Service ClickToRunSvc -ErrorAction Stop
        Write-Host 'Servicio Click-to-Run iniciado.'
    } catch {
        Write-Warning "No se pudo iniciar Click-to-Run: $($_.Exception.Message)"
    }

    try {
        Write-Host 'Solicitando actualización y reparación de Office...'
        Start-Process -FilePath $client -ArgumentList '/update user displaylevel=true forceappshutdown=true' -Wait
    } catch {
        Write-Warning "La reparación no pudo ejecutarse: $($_.Exception.Message)"
    }
}

function Invoke-UninstallCommand {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$Command
    )

    Write-Host "Intentando desinstalar: $DisplayName"
    try {
        $cmd = [Environment]::ExpandEnvironmentVariables($Command.Trim())

        if ($cmd -match '(?i)msiexec(?:\.exe)?\s+(.*)') {
            $args = $Matches[1] -replace '(?i)/I(?=\s*\{)', '/X'
            if ($args -notmatch '(?i)/q') { $args += ' /qn /norestart' }
            Start-Process "$env:SystemRoot\System32\msiexec.exe" -ArgumentList $args -Wait
            return
        }

        $officeClickToRun = "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
        $binaryHealth = Test-BinaryHealth $officeClickToRun
        if ($cmd -match '(?i)(OfficeClickToRun|OfficeC2RClient)' -and -not $binaryHealth.Healthy) {
            Write-Warning 'Se omite el desinstalador Click-to-Run porque el binario falta o su firma no es válida.'
            return
        }

        if ($cmd.StartsWith('"')) {
            $secondQuote = $cmd.IndexOf('"', 1)
            $exe = $cmd.Substring(1, $secondQuote - 1)
            $args = $cmd.Substring($secondQuote + 1).Trim()
        } else {
            $firstSpace = $cmd.IndexOf(' ')
            if ($firstSpace -gt 0) {
                $exe = $cmd.Substring(0, $firstSpace)
                $args = $cmd.Substring($firstSpace + 1)
            } else {
                $exe = $cmd
                $args = ''
            }
        }

        if (Test-Path -LiteralPath $exe) {
            Start-Process -FilePath $exe -ArgumentList $args -Wait
        } else {
            Write-Warning "El desinstalador registrado no existe: $exe"
        }
    } catch {
        Write-Warning "Falló el desinstalador de $DisplayName | $($_.Exception.Message)"
    }
}

function Backup-OfficeRegistry {
    Write-Step 'Copia de seguridad del registro'
    Export-RegKey 'HKLM\SOFTWARE\Microsoft\Office' 'HKLM-Office.reg'
    Export-RegKey 'HKLM\SOFTWARE\WOW6432Node\Microsoft\Office' 'HKLM-WOW6432Node-Office.reg'
    Export-RegKey 'HKLM\SOFTWARE\Microsoft\Office\ClickToRun' 'HKLM-ClickToRun.reg'
    Export-RegKey 'HKCU\Software\Microsoft\Office' 'HKCU-Office.reg'
    Write-Host "Copias guardadas en: $RegistryBackupDir"
}

function New-RestorePointSafe {
    Write-Step 'Punto de restauración'
    try {
        Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
        Checkpoint-Computer -Description 'Antes de limpiar Microsoft Office - GOINTIC' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        Write-Host 'Punto de restauración creado.'
    } catch {
        Write-Warning 'No se pudo crear el punto de restauración. Se continuará con la copia del registro.'
    }
}

function Remove-OfficeComplete {
    Write-Step 'Limpieza completa de Microsoft Office'
    Write-Warning 'Esta operación desinstalará Office del equipo.'

    if (-not $Force) {
        $confirmation = Read-Host 'Escribe ELIMINAR para continuar'
        if ($confirmation -cne 'ELIMINAR') {
            Write-Host 'Operación cancelada.'
            return
        }
    }

    New-RestorePointSafe
    Backup-OfficeRegistry

    Write-Step 'Cerrar procesos de Office'
    @(
        'WINWORD','EXCEL','POWERPNT','OUTLOOK','MSACCESS','ONENOTE',
        'MSPUB','VISIO','WINPROJ','LYNC','OfficeClickToRun',
        'OfficeC2RClient','OfficeC2RCom','AppVShNotify','SDXHelper',
        'integratedoffice','firstrun'
    ) | ForEach-Object {
        Get-Process $_ -ErrorAction SilentlyContinue |
            Stop-Process -Force -ErrorAction SilentlyContinue
    }

    Write-Step 'Desinstaladores registrados'
    foreach ($entry in @(Get-OfficeEntries)) {
        $command = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { $entry.UninstallString }
        if ($command) {
            Invoke-UninstallCommand -DisplayName $entry.DisplayName -Command $command
        }
    }

    Write-Step 'Paquetes Microsoft Store'
    Get-AppxPackage -AllUsers -Name 'Microsoft.Office.Desktop' -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction Stop
                Write-Host "Paquete eliminado: $($_.PackageFullName)"
            } catch {
                Write-Warning "No se pudo eliminar el paquete Appx: $($_.Exception.Message)"
            }
        }

    Write-Step 'Servicios de Office'
    @('ClickToRunSvc','ose','osppsvc') | ForEach-Object {
        $svc = Get-Service $_ -ErrorAction SilentlyContinue
        if ($svc) {
            Stop-Service $_ -Force -ErrorAction SilentlyContinue
            & sc.exe delete $_ | Out-Null
            Write-Host "Servicio marcado para eliminación: $_"
        }
    }

    Write-Step 'Tareas programadas de Office'
    Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object {
            $_.TaskPath -like '\Microsoft\Office\*' -or
            $_.TaskName -match '(?i)^(Office|Microsoft Office)'
        } |
        ForEach-Object {
            try {
                Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -Confirm:$false -ErrorAction Stop
                Write-Host "Tarea eliminada: $($_.TaskPath)$($_.TaskName)"
            } catch {
                Write-Warning "No se pudo eliminar la tarea $($_.TaskName)"
            }
        }

    Write-Step 'Carpetas de Office'
    @(
        "$env:ProgramFiles\Microsoft Office",
        "${env:ProgramFiles(x86)}\Microsoft Office",
        "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun",
        "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun",
        "$env:ProgramData\Microsoft\ClickToRun",
        "$env:ProgramData\Microsoft\Office",
        "$env:SystemDrive\MSOCache",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Office"
    ) | Where-Object { $_ } | ForEach-Object { Remove-PathSafe $_ }

    Write-Step 'Registro de Office'
    @(
        'HKLM:\SOFTWARE\Microsoft\Office',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office',
        'HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeSoftwareProtectionPlatform'
    ) | ForEach-Object { Remove-RegKeySafe $_ }

    foreach ($root in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )) {
        Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
            $item = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if (Test-OfficeName $item.DisplayName) {
                Remove-RegKeySafe $_.PSPath
            }
        }
    }

    if ($PurgeUserSettings) {
        Write-Step 'Configuración del usuario'
        @(
            "$env:LOCALAPPDATA\Microsoft\Office",
            "$env:APPDATA\Microsoft\Office"
        ) | ForEach-Object { Remove-PathSafe $_ }
        Remove-RegKeySafe 'HKCU:\Software\Microsoft\Office'
    } else {
        Write-Host 'Configuración del usuario conservada.'
    }

    if ($RemoveTeamsAddin) {
        Write-Step 'Complemento de reuniones de Teams'
        Get-OfficeEntries | Where-Object { $_.DisplayName -match '(?i)Teams Meeting Add-in' } |
            ForEach-Object {
                $cmd = if ($_.QuietUninstallString) { $_.QuietUninstallString } else { $_.UninstallString }
                if ($cmd) { Invoke-UninstallCommand $_.DisplayName $cmd }
            }
    }

    Write-Step 'Comprobación final'
    $remaining = @(Get-OfficeEntries)
    if ($remaining.Count -eq 0) {
        Write-Host 'No se detectan suites Office registradas.' -ForegroundColor Green
    } else {
        Write-Warning 'Persisten las siguientes entradas:'
        $remaining | ForEach-Object { Write-Warning " - $($_.DisplayName)" }
    }

    Write-Host ''
    Write-Host 'Limpieza finalizada. Reinicia Windows antes de reinstalar Office.' -ForegroundColor Green
}

function Invoke-SystemRepair {
    Write-Step 'Reparación de Windows con DISM y SFC'
    Write-Host 'Ejecutando DISM /RestoreHealth...'
    & DISM.exe /Online /Cleanup-Image /RestoreHealth
    Write-Host ''
    Write-Host 'Ejecutando SFC /scannow...'
    & sfc.exe /scannow
}

function Install-OfficeWithODT {
    Write-Step 'Instalación de Office con Office Deployment Tool'
    Write-Host 'Esta opción usa un setup.exe y un configuration.xml que ya existan en el equipo.'
    Write-Host 'No descarga Office Deployment Tool automáticamente.'
    Write-Host ''

    $setupPath = Read-Host 'Ruta completa de setup.exe'
    $configPath = Read-Host 'Ruta completa de configuration.xml'

    if (-not (Test-Path -LiteralPath $setupPath)) {
        Write-Warning 'No se encuentra setup.exe.'
        return
    }
    if (-not (Test-Path -LiteralPath $configPath)) {
        Write-Warning 'No se encuentra configuration.xml.'
        return
    }

    try {
        Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -Wait
        Write-Host 'El instalador ODT ha terminado.' -ForegroundColor Green
    } catch {
        Write-Warning "Error al ejecutar ODT: $($_.Exception.Message)"
    }
}

function Convert-ToHtmlSafe {
    param([object]$Value)
    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-HtmlReport {
    Write-Step 'Generar informe HTML'

    if ($script:DiagnosticData.Count -eq 0) {
        Get-OfficeDiagnostics | Out-Null
    }

    $entriesHtml = if (@($script:DiagnosticData.OfficeEntries).Count -gt 0) {
        (@($script:DiagnosticData.OfficeEntries) | ForEach-Object {
            "<tr><td>$(Convert-ToHtmlSafe $_.DisplayName)</td><td>$(Convert-ToHtmlSafe $_.DisplayVersion)</td><td>$(Convert-ToHtmlSafe $_.Publisher)</td></tr>"
        }) -join "`n"
    } else {
        '<tr><td colspan="3">No se detectaron productos Office registrados.</td></tr>'
    }

    $wordPaths = if (@($script:DiagnosticData.WinWordPaths).Count -gt 0) {
        (@($script:DiagnosticData.WinWordPaths) | ForEach-Object { "<li>$(Convert-ToHtmlSafe $_)</li>" }) -join "`n"
    } else {
        '<li>No detectado</li>'
    }

    $html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>GOINTIC Office Repair Report</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 32px; color: #1f2937; }
h1 { margin-bottom: 4px; }
h2 { margin-top: 28px; border-bottom: 1px solid #d1d5db; padding-bottom: 6px; }
.badge { display:inline-block; padding:4px 8px; border-radius:6px; background:#e5e7eb; }
table { border-collapse: collapse; width:100%; margin-top:12px; }
th, td { border:1px solid #d1d5db; padding:8px; text-align:left; }
th { background:#f3f4f6; }
.small { color:#6b7280; font-size:12px; }
</style>
</head>
<body>
<h1>GOINTIC Office Repair Tool</h1>
<div class="small">Versión $ToolVersion · Generado: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</div>

<h2>Sistema</h2>
<table>
<tr><th>Equipo</th><td>$(Convert-ToHtmlSafe $script:DiagnosticData.Equipo)</td></tr>
<tr><th>Usuario</th><td>$(Convert-ToHtmlSafe $script:DiagnosticData.Usuario)</td></tr>
<tr><th>Windows</th><td>$(Convert-ToHtmlSafe $script:DiagnosticData.Windows) $(Convert-ToHtmlSafe $script:DiagnosticData.WindowsVersion)</td></tr>
<tr><th>Reinicio pendiente</th><td>$(Convert-ToHtmlSafe $script:DiagnosticData.RebootPending)</td></tr>
</table>

<h2>Click-to-Run</h2>
<table>
<tr><th>Servicio</th><td>$(Convert-ToHtmlSafe $script:DiagnosticData.ClickToRunService)</td></tr>
<tr><th>Carpeta</th><td>$(Convert-ToHtmlSafe $script:DiagnosticData.ClickToRunFolder)</td></tr>
<tr><th>OfficeClickToRun.exe</th><td>Existe: $(Convert-ToHtmlSafe $script:DiagnosticData.OfficeClickToRun.Exists) · Firma: $(Convert-ToHtmlSafe $script:DiagnosticData.OfficeClickToRun.Signature) · Versión: $(Convert-ToHtmlSafe $script:DiagnosticData.OfficeClickToRun.Version)</td></tr>
<tr><th>C2RUI.dll</th><td>Existe: $(Convert-ToHtmlSafe $script:DiagnosticData.C2RUI.Exists) · Firma: $(Convert-ToHtmlSafe $script:DiagnosticData.C2RUI.Signature) · Versión: $(Convert-ToHtmlSafe $script:DiagnosticData.C2RUI.Version)</td></tr>
</table>

<h2>Productos Office detectados</h2>
<table>
<tr><th>Producto</th><th>Versión</th><th>Editor</th></tr>
$entriesHtml
</table>

<h2>WINWORD.EXE</h2>
<ul>
$wordPaths
</ul>

<h2>Archivos del expediente</h2>
<ul>
<li>Log: $(Convert-ToHtmlSafe $LogFile)</li>
<li>Copias de registro: $(Convert-ToHtmlSafe $RegistryBackupDir)</li>
</ul>

<p class="small">Herramienta independiente. No es un producto oficial de Microsoft.</p>
</body>
</html>
"@

    Set-Content -Path $ReportFile -Value $html -Encoding UTF8
    Write-Host "Informe generado: $ReportFile" -ForegroundColor Green
}

function Show-Menu {
    do {
        Write-Title
        Write-Host '[1] Diagnóstico de Office'
        Write-Host '[2] Reparación rápida de Click-to-Run'
        Write-Host '[3] Limpieza completa de Office'
        Write-Host '[4] Ejecutar DISM y SFC'
        Write-Host '[5] Reinstalar Office mediante ODT existente'
        Write-Host '[6] Generar informe HTML'
        Write-Host '[7] Salir'
        Write-Host ''

        $choice = Read-Host 'Selecciona una opción'
        switch ($choice) {
            '1' { Get-OfficeDiagnostics | Out-Null; Wait-Key }
            '2' { Repair-ClickToRun; Wait-Key }
            '3' { Remove-OfficeComplete; New-HtmlReport; Wait-Key }
            '4' { Invoke-SystemRepair; Wait-Key }
            '5' { Install-OfficeWithODT; Wait-Key }
            '6' { New-HtmlReport; Wait-Key }
            '7' { return }
            default { Write-Warning 'Opción no válida.'; Start-Sleep 1 }
        }
    } while ($true)
}

try {
    if ($Silent) {
        Get-OfficeDiagnostics | Out-Null
        New-HtmlReport
    } else {
        Show-Menu
    }
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
    Write-Host ''
    Write-Host "Archivos generados en: $BaseDir" -ForegroundColor Yellow
}

