<# :
@echo off
setlocal
fsutil dirty query %systemdrive% >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%0' -Verb RunAs"
    exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ([System.IO.File]::ReadAllText('%~f0'))"
pause
exit /b
#>

# ==============================================================
#   ENDPOINT MAINTENANCE TOOL v1.0
#   Autor: Lucca Oliveira | github.com/luccaolvr
#
#   Modos de execucao:
#     [1] Rapido  — Coleta de dados + alertas (~3-5 min)
#     [2] Completo — Rapido + limpeza + otimizacao (~20-45 min)
#
#   CONFIGURACAO NECESSARIA (antes de usar):
#     Edite o bloco "CONFIGURACAO" abaixo com os dados do seu ambiente.
# ==============================================================


# ==============================================================
# CONFIGURACAO — edite antes de usar
# ==============================================================
$CONFIG = @{
    # Caminho de rede onde os arquivos serao salvos
    # Exemplo: "\\servidor\compartilhamento\pasta"
    NetworkShare   = "\\SEU_SERVIDOR\SEU_COMPARTILHAMENTO"

    # Subpasta dentro do compartilhamento para salvar os arquivos
    SubFolder      = "Preventiva"

    # Nome do arquivo Excel de inventario (deve existir na pasta acima)
    ExcelFile      = "Inventario.xlsx"

    # Nome da aba do Excel onde os dados serao gravados
    ExcelSheet     = "Inventario"

    # Credencial de rede (usuario com acesso ao compartilhamento)
    # Deixe em branco para usar a sessao atual do Windows
    NetUser        = ""   # Exemplo: "DOMINIO\usuario"
    NetPass        = ""   # Exemplo: "senha"
}
# ==============================================================


# --------------------------------------------------------------
# FUNCOES UTILITARIAS
# --------------------------------------------------------------

# Barra de progresso visual no terminal
function Show-Progress($Message, $Percent) {
    $Width   = 25
    $Filled  = [Math]::Floor($Percent / 100 * $Width)
    $Bar     = "[" + ("#" * $Filled) + ("-" * ($Width - $Filled)) + "]"
    Write-Host "`r$Message $Bar $Percent%" -NoNewline -ForegroundColor Cyan
    if ($Percent -eq 100) { Write-Host " [OK]" -ForegroundColor Green }
}

# Pergunta Sim/Nao e retorna $true/$false
function Ask-YesNo($Question) {
    do {
        Write-Host $Question -ForegroundColor Yellow -NoNewline
        Write-Host " (S/N): " -NoNewline
        $Answer = (Read-Host).Trim().ToUpper()
    } while ($Answer -ne "S" -and $Answer -ne "N")
    return $Answer -eq "S"
}

# Pergunta texto livre e retorna o valor
function Ask-Text($Question) {
    Write-Host $Question -ForegroundColor Yellow -NoNewline
    Write-Host ": " -NoNewline
    return (Read-Host).Trim()
}

# Grava valor texto em celula Excel
function Set-ExcelStr($Cell, $Value) { $Cell.Value2 = [string]$Value }

# Grava valor numerico em celula Excel
function Set-ExcelNum($Cell, $Value) { $Cell.Value2 = [double]$Value }

$ErrorActionPreference = "SilentlyContinue"


# --------------------------------------------------------------
# SELECAO DE MODO
# --------------------------------------------------------------
Clear-Host
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "   ENDPOINT MAINTENANCE TOOL v1.0                       " -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Selecione o modo de execucao:" -ForegroundColor White
Write-Host ""
Write-Host "  [1] Rapido   — Diagnostico e inventario (~3-5 min)" -ForegroundColor Green
Write-Host "  [2] Completo — Diagnostico + limpeza + otimizacao (~20-45 min)" -ForegroundColor Yellow
Write-Host ""

do {
    Write-Host "Modo: " -NoNewline -ForegroundColor Cyan
    $ModeInput = (Read-Host).Trim()
} while ($ModeInput -ne "1" -and $ModeInput -ne "2")

$QuickMode = ($ModeInput -eq "1")
$ModeName  = if ($QuickMode) { "RAPIDO" } else { "COMPLETO" }

Write-Host ""
Write-Host "  Modo selecionado: $ModeName" -ForegroundColor Magenta
Write-Host ""


# --------------------------------------------------------------
# COLETA MANUAL DE DADOS DO ATENDIMENTO
# --------------------------------------------------------------
Write-Host "--- DADOS DO ATENDIMENTO ---" -ForegroundColor Magenta
Write-Host ""

# Nome do usuario
do {
    $UserName = Ask-Text "Nome completo do usuario"
} while ([string]::IsNullOrWhiteSpace($UserName))

# Setor
do {
    $Department = Ask-Text "Setor"
} while ([string]::IsNullOrWhiteSpace($Department))

# Tipo de ativo
$IsNotebook = Ask-YesNo "O equipamento e um Notebook?"
$AssetType  = if ($IsNotebook) { "Notebook" } else { "Desktop" }

# Monitor
$HasMonitor    = Ask-YesNo "Possui monitor?"
$MonitorModel  = ""
$MonitorAsset  = ""
$MonitorInputs = ""
if ($HasMonitor) {
    $MonitorModel  = Ask-Text "Modelo do monitor"
    $MonitorAsset  = Ask-Text "Patrimonio do monitor (deixe em branco se nao souber)"
    $MonitorInputs = Ask-Text "Entradas do monitor (ex: HDMI, VGA, DP)"
}
$MonitorCol = if ($HasMonitor) { $MonitorModel } else { "Nao possui" }

# Teclado e mouse
$HasPeripherals = Ask-YesNo "Possui teclado e mouse?"
$KeyboardModel  = ""
$MouseModel     = ""
if ($HasPeripherals) {
    $KeyboardModel = Ask-Text "Modelo do teclado"
    $MouseModel    = Ask-Text "Modelo do mouse"
}
$KeyboardCol = if ($HasPeripherals) { $KeyboardModel } else { "Nao possui" }
$MouseCol    = if ($HasPeripherals) { $MouseModel    } else { "Nao possui" }

Write-Host ""


# --------------------------------------------------------------
# MODO COMPLETO — Otimizacao e limpeza
# (executado antes do diagnostico para refletir estado pos-limpeza)
# --------------------------------------------------------------
if (-not $QuickMode) {

    Write-Host "=========================================================" -ForegroundColor Yellow
    Write-Host "   MODO COMPLETO — Iniciando limpeza e otimizacao       " -ForegroundColor Yellow
    Write-Host "=========================================================" -ForegroundColor Yellow
    Write-Host ""

    # --- Remocao de apps desnecessarios (bloatware) ---
    Show-Progress "Removendo bloatware         " 10
    $Bloatware = @(
        "*Solitaire*", "*Skype*", "*Xbox*",
        "*YourPhone*", "*ZuneVideo*", "*BingNews*"
    )
    foreach ($App in $Bloatware) {
        Get-AppxPackage -AllUsers -Name $App |
            Remove-AppxPackage -AllUsers
        Get-AppxProvisionedPackage -Online |
            Where-Object { $_.DisplayName -like $App } |
            Remove-AppxProvisionedPackage -Online
    }
    Show-Progress "Removendo bloatware         " 100

    # --- Desativa apps em segundo plano ---
    Show-Progress "Otimizando inicializacao    " 40
    Set-ItemProperty `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications" `
        -Name "GlobalUserDisabled" `
        -Value 1
    Show-Progress "Otimizando inicializacao    " 100

    # --- Cria usuario de suporte local (se nao existir) ---
    # NOTA: defina a senha do usuario de suporte abaixo antes de usar
    Show-Progress "Verificando usuario suporte " 60
    $SupportUser = "suporte.ti"       # <-- altere para o nome desejado
    $SupportPass = "DEFINA_UMA_SENHA" # <-- altere para a senha desejada
    if (-not (Get-LocalUser -Name $SupportUser -ErrorAction SilentlyContinue)) {
        $SecurePass = ConvertTo-SecureString $SupportPass -AsPlainText -Force
        New-LocalUser -Name $SupportUser -Password $SecurePass -FullName "Suporte TI" | Out-Null
        Add-LocalGroupMember -Group "S-1-5-32-544" -Member $SupportUser
    }
    Show-Progress "Verificando usuario suporte " 100

    # --- Limpeza de temporarios e logs ---
    Show-Progress "Limpando temporarios/logs   " 80
    $CleanPaths = @(
        "C:\Windows\Temp\*",
        "$env:LOCALAPPDATA\Temp\*",
        "C:\Windows\Prefetch\*"
    )
    foreach ($Path in $CleanPaths) {
        Remove-Item -Path $Path -Recurse -Force
    }
    foreach ($Log in @("Application", "System")) {
        [System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog($Log)
    }
    Show-Progress "Limpando temporarios/logs   " 100

    Write-Host ""
    Write-Host "[OK] Limpeza e otimizacao concluidas." -ForegroundColor Green
    Write-Host ""
}


# --------------------------------------------------------------
# COLETA DE DIAGNOSTICO E HARDWARE
# --------------------------------------------------------------
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "   Coletando diagnostico do sistema...                  " -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

Show-Progress "Coletando hardware/sistema  " 20

$SysInfo  = Get-CimInstance Win32_ComputerSystem
$OSInfo   = Get-CimInstance Win32_OperatingSystem
$DiskInfo = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$MemInfo  = Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1
$AVInfo   = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct |
            Where-Object { $_.displayName -like "*Bitdefender*" }

# IP da maquina (ignora APIPA e loopback)
$IPAddress = (
    Get-NetIPAddress -AddressFamily IPv4 |
    Where-Object {
        $_.IPAddress -notmatch "^169" -and
        $_.InterfaceAlias -notlike "*Loopback*"
    }
).IPAddress -join " / "

# Usuario logado e e-mail no AD
$ADUsername = $env:USERNAME
$ADEmail    = ""
try {
    $ADResult = ([adsisearcher]"(sAMAccountName=$ADUsername)").FindOne()
    $ADEmail  = $ADResult.Properties["mail"][0]
} catch {}

Show-Progress "Coletando hardware/sistema  " 50

# Tipo de memoria RAM
$RAMType = switch ($MemInfo.SMBIOSMemoryType) {
    20 { "DDR"  }
    21 { "DDR2" }
    24 { "DDR3" }
    26 { "DDR4" }
    34 { "DDR5" }
    0  { if ($MemInfo.ConfiguredClockSpeed -gt 4000) { "DDR5" } else { "DDR4" } }
    default { "DDRx" }
}

# Metricas de disco
$DiskGB       = [Math]::Round($DiskInfo.Size / 1GB, 0)
$DiskFreeGB   = [Math]::Round($DiskInfo.FreeSpace / 1GB, 0)
$DiskFreePct  = [Math]::Round(($DiskInfo.FreeSpace / $DiskInfo.Size) * 100, 0)
$DiskUsedPct  = 100 - $DiskFreePct
$DiskType     = (Get-PhysicalDisk | Select-Object -First 1).MediaType

# Metricas de RAM
$TotalRAM_GB  = [Math]::Round($SysInfo.TotalPhysicalMemory / 1GB, 0)
$TotalRAM_MB  = $SysInfo.TotalPhysicalMemory / 1MB
$FreeRAM_MB   = $OSInfo.FreePhysicalMemory / 1KB
$UsedRAM_MB   = $TotalRAM_MB - $FreeRAM_MB
$RAMUsedPct   = [Math]::Round(($UsedRAM_MB / $TotalRAM_MB) * 100, 0)

Show-Progress "Coletando hardware/sistema  " 70

# Drivers com erro (filtra virtuais e genericos)
$IgnoreClasses = @("Processor", "Computer", "SCSIAdapter")
$IgnoreNames   = @(
    "*Microsoft Virtual*", "*Remote Desktop*", "*VMware*",
    "*VirtualBox*", "*Generic volume*", "*Composite Bus*", "*NDIS Virtual*"
)
$DriversWithError = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.ConfigManagerErrorCode -ne 0 -and
    $_.ConfigManagerErrorCode -ne 22 -and
    $_.PNPClass -notin $IgnoreClasses -and
    -not ($IgnoreNames | Where-Object { $_ -like $_.Name })
}

# Uptime
$UptimeDays = [Math]::Round(((Get-Date) - $OSInfo.LastBootUpTime).TotalDays, 1)
$UptimeText = if ($UptimeDays -gt 30) { "ALTO ($UptimeDays dias)" } else { "OK ($UptimeDays dias)" }
$UptimeColor = if ($UptimeDays -gt 30) { "Yellow" } else { "Green" }

# Saude do disco (SMART)
$SmartFailed = $false
$SmartText   = "Nao disponivel"
try {
    $SmartData   = Get-PhysicalDisk | Select-Object -First 1 | Get-StorageReliabilityCounter
    $SmartFailed = (
        $SmartData.ReadErrorsTotal  -gt 50 -or
        $SmartData.WriteErrorsTotal -gt 50 -or
        ($SmartData.Temperature -gt 55 -and $SmartData.Temperature -gt 0)
    )
    $SmartText = if ($SmartFailed) { "ALERTA" } else { "OK" }
    if ($SmartData.Temperature -gt 0) { $SmartText += " | Temp: $($SmartData.Temperature)C" }
} catch {}

# Top 5 processos por CPU e RAM
$TopByCPU = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5 |
    ForEach-Object { "  $($_.ProcessName.PadRight(25)) CPU: $([Math]::Round($_.CPU, 1))s" }

$TopByRAM = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 5 |
    ForEach-Object {
        $MB  = [Math]::Round($_.WorkingSet64 / 1MB, 0)
        $Pct = [Math]::Round(($_.WorkingSet64 / 1MB) / $TotalRAM_MB * 100, 1)
        "  $($_.ProcessName.PadRight(25)) RAM: $MB MB ($Pct%)"
    }

# Processos consumindo mais de 80% da RAM
$HeavyProcesses = Get-Process |
    Where-Object { (($_.WorkingSet64 / 1MB) / $TotalRAM_MB * 100) -gt 80 } |
    ForEach-Object { $_.ProcessName }

# Reinicializacao pendente do Windows Update
$RebootPending = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"

# Eventos criticos nas ultimas 24h
$CriticalEvents = 0
try {
    $CriticalEvents = (
        Get-WinEvent -FilterHashtable @{
            LogName   = "System"
            Level     = 1, 2
            StartTime = (Get-Date).AddDays(-1)
        } -ErrorAction SilentlyContinue
    ).Count
} catch {}

# Ultima atualizacao instalada
$LastUpdate = (Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1).InstalledOn

Show-Progress "Coletando hardware/sistema  " 100


# --------------------------------------------------------------
# EXIBICAO DO STATUS DE SAUDE
# --------------------------------------------------------------
Write-Host ""
Write-Host "--- STATUS DE SAUDE ---" -ForegroundColor Yellow

# Antivirus
$AVStatus    = if ($AVInfo) { "OPERANTE" } else { "ALERTA" }
$AVColor     = if ($AVInfo) { "Green"    } else { "Red"    }
Write-Host "Antivirus:    " -NoNewline
Write-Host $AVStatus -ForegroundColor $AVColor

# Espaco em disco
$DiskColor = if ($DiskUsedPct -ge 80) { "Red" } else { "Green" }
Write-Host "Espaco Livre: " -NoNewline
Write-Host "$DiskFreePct% ($DiskFreeGB GB livres)" -ForegroundColor $DiskColor
if ($DiskUsedPct -ge 85) {
    Write-Host "              !! CRITICO: $DiskUsedPct% em uso — necessario limpeza urgente" -ForegroundColor Red
} elseif ($DiskUsedPct -ge 80) {
    Write-Host "              !! ATENCAO: $DiskUsedPct% em uso — recomendado limpeza" -ForegroundColor Yellow
}

# Uso de RAM
$RAMColor = if ($RAMUsedPct -ge 85) { "Red" } else { "Green" }
Write-Host "Uso de RAM:   " -NoNewline
Write-Host "$RAMUsedPct% ($TotalRAM_GB GB total)" -ForegroundColor $RAMColor
if ($RAMUsedPct -ge 85) {
    Write-Host "              !! CRITICO: $RAMUsedPct% de RAM em uso — considerar upgrade" -ForegroundColor Red
}

# Reinicializacao
$RebootStatus = if ($RebootPending) { "PENDENTE" } else { "OK" }
$RebootColor  = if ($RebootPending) { "Yellow"   } else { "Green" }
Write-Host "Reboot:       " -NoNewline
Write-Host $RebootStatus -ForegroundColor $RebootColor

# Drivers
$DriverStatus = if ($DriversWithError) { "ERRO ($($DriversWithError.Count) dispositivo(s))" } else { "OK" }
$DriverColor  = if ($DriversWithError) { "Red" } else { "Green" }
Write-Host "Drivers:      " -NoNewline
Write-Host $DriverStatus -ForegroundColor $DriverColor
if ($DriversWithError) {
    foreach ($Dev in $DriversWithError) {
        Write-Host "              >> $($Dev.Name) [Erro $($Dev.ConfigManagerErrorCode)]" -ForegroundColor DarkYellow
    }
}

# Uptime e SMART
Write-Host "Uptime:       " -NoNewline
Write-Host $UptimeText -ForegroundColor $UptimeColor

$SmartColor = if ($SmartFailed) { "Red" } else { "Green" }
Write-Host "Disco SMART:  " -NoNewline
Write-Host $SmartText -ForegroundColor $SmartColor


# --------------------------------------------------------------
# LISTA DE ALERTAS
# --------------------------------------------------------------
$Alerts = @()
if (-not $AVInfo)           { $Alerts += "!! Antivirus nao detectado ou inativo" }
if ($DiskUsedPct -ge 85)    { $Alerts += "!! Disco com $DiskUsedPct% de uso — limpeza urgente necessaria" }
elseif ($DiskUsedPct -ge 80){ $Alerts += "!! Disco com $DiskUsedPct% de uso — recomendado limpeza" }
if ($RAMUsedPct -ge 85)     { $Alerts += "!! RAM com $RAMUsedPct% de uso ($TotalRAM_GB GB) — considerar upgrade" }
if ($DriversWithError)      { $Alerts += "!! $($DriversWithError.Count) driver(s) com erro detectado(s)" }
if ($SmartFailed)           { $Alerts += "!! Falha SMART detectada no disco" }
if ($UptimeDays -gt 30)     { $Alerts += "!! Maquina sem reiniciar ha $UptimeDays dias" }
if ($RebootPending)         { $Alerts += "!! Reinicializacao pendente (Windows Update)" }
if ($HeavyProcesses)        { $Alerts += "!! Processo(s) acima de 80% de RAM: $($HeavyProcesses -join ', ')" }
if ($CriticalEvents -gt 0)  { $Alerts += "!! $CriticalEvents evento(s) critico(s) no log do sistema (ultimas 24h)" }

Write-Host ""
if ($Alerts.Count -gt 0) {
    Write-Host "=========================================================" -ForegroundColor Red
    Write-Host "   ATENCAO — $($Alerts.Count) ALERTA(S) ENCONTRADO(S)   " -ForegroundColor Red
    Write-Host "=========================================================" -ForegroundColor Red
    foreach ($Alert in $Alerts) { Write-Host "  $Alert" -ForegroundColor Yellow }
    Write-Host "=========================================================" -ForegroundColor Red
} else {
    Write-Host "[OK] Nenhum problema critico encontrado." -ForegroundColor Green
}


# --------------------------------------------------------------
# TOP PROCESSOS
# --------------------------------------------------------------
Write-Host ""
Write-Host "--- TOP 5 PROCESSOS POR CPU ---" -ForegroundColor Yellow
$TopByCPU | ForEach-Object { Write-Host $_ -ForegroundColor White }

Write-Host ""
Write-Host "--- TOP 5 PROCESSOS POR RAM ---" -ForegroundColor Yellow
$TopByRAM | ForEach-Object { Write-Host $_ -ForegroundColor White }


# --------------------------------------------------------------
# BLOCO DE DADOS PARA MOVIDESK / ITSM
# --------------------------------------------------------------
$Separator   = "---------------------------------------------------------"
$DriverDetail = if ($DriversWithError) {
    ($DriversWithError | ForEach-Object { "$($_.Name) [Erro $($_.ConfigManagerErrorCode)]" }) -join "; "
} else { "OK" }

$ITSMBlock = @"
$Separator
DADOS PARA REGISTRO NO ITSM
$Separator
Hostname:      $($SysInfo.Name)
Usuario:       $UserName ($ADUsername)
Setor:         $Department
Tipo Ativo:    $AssetType
IP:            $IPAddress
Fabricante:    $($SysInfo.Manufacturer)
Modelo:        $($SysInfo.Model)
Serial:        $((Get-CimInstance Win32_Bios).SerialNumber)
Processador:   $((Get-CimInstance Win32_Processor).Name)
RAM:           $TotalRAM_GB GB $RAMType ($RAMUsedPct% em uso)
Armazenamento: $DiskGB GB ($DiskType) | Livre: $DiskFreeGB GB ($DiskFreePct%) | Uso: $DiskUsedPct%
Sistema Op.:   $($OSInfo.Caption)
Monitor:       $MonitorCol
Teclado:       $KeyboardCol
Mouse:         $MouseCol
Uptime:        $UptimeText
Disco SMART:   $SmartText
Ultima Att.:   $LastUpdate
Antivirus:     $AVStatus
Drivers:       $DriverDetail
Reboot:        $RebootStatus
"@

Write-Host ""
Write-Host $ITSMBlock -ForegroundColor White
$ITSMBlock | clip
Write-Host ""
Write-Host "[OK] Dados copiados para a area de transferencia." -ForegroundColor Green


# --------------------------------------------------------------
# SALVAR RELATORIO TXT NA REDE
# --------------------------------------------------------------
$AlertsText  = if ($Alerts.Count -gt 0) {
    "=========================================================`r`n" +
    "   ATENCAO — $($Alerts.Count) ALERTA(S) ENCONTRADO(S)`r`n" +
    "=========================================================`r`n" +
    (($Alerts | ForEach-Object { "  $_" }) -join "`r`n") + "`r`n" +
    "========================================================="
} else { "[OK] Nenhum problema critico encontrado." }

$ReportContent = @"
=========================================================
   ENDPOINT MAINTENANCE TOOL v1.0
   Modo:   $ModeName
   Data:   $(Get-Date -Format "dd/MM/yyyy HH:mm")
   Tecnico: $env:USERNAME
=========================================================

$AlertsText

--- TOP 5 PROCESSOS POR CPU ---
$($TopByCPU -join "`r`n")

--- TOP 5 PROCESSOS POR RAM ---
$($TopByRAM -join "`r`n")

$ITSMBlock
=========================================================
"@

$ReportName = "Maintenance_$($UserName -replace ' ','_')_$(Get-Date -Format 'yyyy-MM-dd_HHmm').txt"

try {
    # Monta drive de rede
    if (Get-PSDrive -Name "MAINT" -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name "MAINT" -Force
    }

    if ($CONFIG.NetUser -ne "") {
        $Credential = New-Object System.Management.Automation.PSCredential(
            $CONFIG.NetUser,
            (ConvertTo-SecureString $CONFIG.NetPass -AsPlainText -Force)
        )
        New-PSDrive -Name "MAINT" -PSProvider FileSystem `
            -Root $CONFIG.NetworkShare -Credential $Credential -Persist:$false | Out-Null
    } else {
        New-PSDrive -Name "MAINT" -PSProvider FileSystem `
            -Root $CONFIG.NetworkShare -Persist:$false | Out-Null
    }

    $ReportPath = "MAINT:\$($CONFIG.SubFolder)\$ReportName"

    if (Test-Path "MAINT:\$($CONFIG.SubFolder)") {
        $ReportContent | Out-File -FilePath $ReportPath -Encoding UTF8 -Force
        Write-Host "[LOG] Relatorio salvo: $ReportName" -ForegroundColor Green
    } else {
        Write-Host "[AVISO] Subpasta nao encontrada: $($CONFIG.SubFolder)" -ForegroundColor Yellow
    }

} catch {
    Write-Host "[ERRO] Nao foi possivel salvar o relatorio: $_" -ForegroundColor Red
} finally {
    if (Get-PSDrive -Name "MAINT" -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name "MAINT" -Force
    }
}


# --------------------------------------------------------------
# GRAVAR NA PLANILHA XLSX (INVENTARIO)
# --------------------------------------------------------------
Write-Host ""
Write-Host "[XLSX] Gravando na planilha de inventario..." -ForegroundColor Cyan

$Excel     = $null
$Workbook  = $null
$Worksheet = $null
$LocalXlsx = $null

try {
    # Remove arquivos temporarios de execucoes anteriores
    Get-ChildItem "$env:TEMP" -Filter "Maint_*.xlsx" -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue

    # Monta drive de rede (reutiliza ou recria)
    if (-not (Get-PSDrive -Name "MAINT" -ErrorAction SilentlyContinue)) {
        if ($CONFIG.NetUser -ne "") {
            $Credential = New-Object System.Management.Automation.PSCredential(
                $CONFIG.NetUser,
                (ConvertTo-SecureString $CONFIG.NetPass -AsPlainText -Force)
            )
            New-PSDrive -Name "MAINT" -PSProvider FileSystem `
                -Root $CONFIG.NetworkShare -Credential $Credential -Persist:$false | Out-Null
        } else {
            New-PSDrive -Name "MAINT" -PSProvider FileSystem `
                -Root $CONFIG.NetworkShare -Persist:$false | Out-Null
        }
    }

    $RemoteXlsx = "MAINT:\$($CONFIG.SubFolder)\$($CONFIG.ExcelFile)"
    $LocalXlsx  = "$env:TEMP\Maint_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$($env:COMPUTERNAME).xlsx"

    # Copia planilha para local temporario e abre via COM
    Copy-Item -Path $RemoteXlsx -Destination $LocalXlsx -Force

    $Excel            = New-Object -ComObject Excel.Application
    $Excel.Visible    = $false
    $Excel.DisplayAlerts = $false
    $Workbook         = $Excel.Workbooks.Open($LocalXlsx)
    $Worksheet        = $Workbook.Sheets.Item($CONFIG.ExcelSheet)

    # Encontra proxima linha vazia (assume cabecalho nas primeiras linhas)
    $Row = 3
    do { $Row++ } while (
        $Worksheet.Cells($Row, 1).Value2 -ne $null -and
        $Worksheet.Cells($Row, 1).Value2 -ne ""
    )

    # Grava dados na planilha
    Set-ExcelStr $Worksheet.Cells($Row,  1)  (Get-Date -Format "dd/MM/yyyy")
    Set-ExcelStr $Worksheet.Cells($Row,  2)  $Department
    Set-ExcelStr $Worksheet.Cells($Row,  3)  $UserName
    Set-ExcelStr $Worksheet.Cells($Row,  4)  $ADUsername
    Set-ExcelStr $Worksheet.Cells($Row,  5)  $SysInfo.Name
    Set-ExcelStr $Worksheet.Cells($Row,  6)  $ADEmail
    Set-ExcelStr $Worksheet.Cells($Row,  7)  (if ($ADEmail -like "*@*") { "Corporativo" } else { "" })
    Set-ExcelStr $Worksheet.Cells($Row,  8)  $AssetType
    Set-ExcelStr $Worksheet.Cells($Row,  9)  $IPAddress
    Set-ExcelStr $Worksheet.Cells($Row, 10)  $SysInfo.Manufacturer
    Set-ExcelStr $Worksheet.Cells($Row, 11)  $SysInfo.Model
    Set-ExcelStr $Worksheet.Cells($Row, 12)  (Get-CimInstance Win32_Bios).SerialNumber
    Set-ExcelStr $Worksheet.Cells($Row, 13)  (Get-CimInstance Win32_Processor).Name
    Set-ExcelNum $Worksheet.Cells($Row, 14)  $TotalRAM_GB
    Set-ExcelStr $Worksheet.Cells($Row, 15)  $RAMType
    Set-ExcelNum $Worksheet.Cells($Row, 16)  $DiskGB
    Set-ExcelStr $Worksheet.Cells($Row, 17)  ([string]$DiskType)
    Set-ExcelNum $Worksheet.Cells($Row, 18)  $DiskFreeGB
    Set-ExcelNum $Worksheet.Cells($Row, 19)  $DiskFreePct
    Set-ExcelStr $Worksheet.Cells($Row, 20)  $OSInfo.Caption
    Set-ExcelStr $Worksheet.Cells($Row, 21)  $MonitorCol
    Set-ExcelStr $Worksheet.Cells($Row, 22)  $MonitorAsset
    Set-ExcelStr $Worksheet.Cells($Row, 23)  $MonitorInputs
    Set-ExcelStr $Worksheet.Cells($Row, 24)  $MouseCol
    Set-ExcelStr $Worksheet.Cells($Row, 25)  $KeyboardCol
    Set-ExcelStr $Worksheet.Cells($Row, 26)  (if ($IsNotebook) { "Sim" } else { "Nao" })
    Set-ExcelNum $Worksheet.Cells($Row, 27)  $UptimeDays
    Set-ExcelStr $Worksheet.Cells($Row, 28)  $AVStatus
    Set-ExcelStr $Worksheet.Cells($Row, 29)  $SmartText
    Set-ExcelNum $Worksheet.Cells($Row, 30)  $CriticalEvents
    Set-ExcelStr $Worksheet.Cells($Row, 31)  (if ($RebootPending) { "Pendente" } else { "OK" })
    Set-ExcelStr $Worksheet.Cells($Row, 32)  $DriverDetail
    Set-ExcelNum $Worksheet.Cells($Row, 33)  $Alerts.Count
    Set-ExcelStr $Worksheet.Cells($Row, 34)  ($Alerts -join " | ")
    Set-ExcelStr $Worksheet.Cells($Row, 35)  $ModeName

    # Salva localmente e copia de volta para o servidor
    $Workbook.Save()
    $Workbook.Close($false)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)  | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)     | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Copy-Item -Path $LocalXlsx -Destination $RemoteXlsx -Force
    Write-Host "[XLSX] Planilha atualizada com sucesso!" -ForegroundColor Green

} catch {
    Write-Host "[ERRO] Nao foi possivel gravar na planilha: $_" -ForegroundColor Red
} finally {
    # Garante liberacao do COM mesmo em caso de falha
    if ($Workbook) { try { $Workbook.Close($false) } catch {} }
    if ($Excel)    {
        try {
            $Excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
        } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($LocalXlsx -and (Test-Path $LocalXlsx)) {
        Remove-Item -Path $LocalXlsx -Force -ErrorAction SilentlyContinue
    }
    if (Get-PSDrive -Name "MAINT" -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name "MAINT" -Force
    }
}


# --------------------------------------------------------------
# FINALIZACAO
# --------------------------------------------------------------
Write-Host ""
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "   CONCLUIDO — Modo: $ModeName                          " -ForegroundColor Cyan
Write-Host "   Dados copiados para area de transferencia.           " -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
