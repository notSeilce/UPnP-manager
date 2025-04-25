# miniUPnP Manager by Seilce
# Требуются права администратора
#Requires -RunAsAdministrator

# Проверка прав администратора и запрос на их получение
if (-not [System.Security.Principal.WindowsIdentity]::GetCurrent().IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Запуск с правами администратора..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command & { $($myInvocation.MyCommand.Definition) }" -Verb RunAs
    exit
}

# Установка кодировки для корректного отображения русского языка
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# Установка стандартной темы консоли для PowerShell

# Переход в нужный раздел реестра
Set-Location HKCU:\Console

# Создание раздела профиля для powershell.exe
$psProfileKey = '.\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe'
if (-not (Test-Path $psProfileKey)) {
    New-Item $psProfileKey | Out-Null
}

Set-Location $psProfileKey
$size = $host.UI.RawUI.WindowSize
$size.Width = 120
$size.Height = 30
$host.UI.RawUI.WindowSize = $size
# Установка шрифта Consolas 16
New-ItemProperty -Path . -Name FaceName   -Value "Consolas"     -PropertyType String -Force
New-ItemProperty -Path . -Name FontFamily -Value 0x00000036     -PropertyType DWord  -Force
New-ItemProperty -Path . -Name FontSize   -Value 0x00100000     -PropertyType DWord  -Force  # 0x0010 = 16 dec
New-ItemProperty -Path . -Name FontWeight -Value 0x00000190     -PropertyType DWord  -Force  # 400 = normal
Set-ConsoleFont 16
# Возвращаемся обратно
Set-Location $env:USERPROFILE

# Устанавливаем визуальные параметры консоли
(Get-Host).UI.RawUI.ForegroundColor = "White"
(Get-Host).UI.RawUI.BackgroundColor = "Black"
(Get-Host).UI.RawUI.CursorSize = 10
(Get-Host).UI.RawUI.WindowTitle = "miniUPnP Manager by Seilce"
Clear-Host

# Путь к файлу с сохраненными портами
$PORTS_FILE = "C:\Windows\System32\nazzy_ports.txt"

# Функция для создания красивого заголовка
function Show-Header {
    param (
        [string]$Title
    )
    $width = $host.UI.RawUI.WindowSize.Width
    Write-Host "`n"
    Write-Host ($Title) -ForegroundColor Yellow
    Write-Host "`n"
}

# Функция для создания меню
function Show-Menu {
    param (
        [string]$Title = 'miniUPnP Manager'
    )
    Clear-Host
    Show-Header $Title
    
    Write-Host "[ 1 ]" -ForegroundColor Green -NoNewline
    Write-Host " Использовать стандартные порты" -ForegroundColor White
    Write-Host "     ├─ 25565 TCP/UDP (Minecraft)" -ForegroundColor Gray
    Write-Host "     ├─ 7656 UDP (ModularVoice)" -ForegroundColor Gray
    Write-Host "     └─ 24454 UDP (SimpleVoice)" -ForegroundColor Gray
    Write-Host
    Write-Host "[ 2 ]" -ForegroundColor Green -NoNewline
    Write-Host " Настроить свои порты" -ForegroundColor White
    Write-Host "[ 3 ]" -ForegroundColor Green -NoNewline
    Write-Host " Проверить статус подключения" -ForegroundColor White
    Write-Host "[ 4 ]" -ForegroundColor Green -NoNewline
    Write-Host " Показать текущие правила" -ForegroundColor White
    Write-Host "[ 5 ]" -ForegroundColor Green -NoNewline
    Write-Host " Удалить все правила" -ForegroundColor White
    Write-Host "[ 6 ]" -ForegroundColor Red -NoNewline
    Write-Host " Выход" -ForegroundColor White
    Write-Host
}

# Функция для проверки наличия upnpc
function Test-Upnpc {
    $upnpcPath = "C:\Windows\System32\upnpc-static.exe"
    if (-not (Test-Path $upnpcPath)) {
        Write-Host "upnpc не установлен. Начинаем установку..." -ForegroundColor Yellow
        Install-Upnpc
    }
}

# Функция установки upnpc
function Install-Upnpc {
    try {
        Write-Host "Скачивание upnpc..." -ForegroundColor Cyan
        $url = "http://miniupnp.free.fr/files/upnpc-exe-win32-20220515.zip"
        $zipPath = Join-Path $env:TEMP "upnpc.zip"
        $extractPath = Join-Path $env:TEMP "temp_upnpc"
        
        # Скачивание файла
        Invoke-WebRequest -Uri $url -OutFile $zipPath
        
        # Создание временной директории и распаковка
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        
        # Копирование файлов
        Write-Host "Копирование файлов в System32..." -ForegroundColor Cyan
        Copy-Item "$extractPath\upnpc-static.exe" "C:\Windows\System32\" -Force
        Copy-Item "$extractPath\upnpc-shared.exe" "C:\Windows\System32\" -Force
        Copy-Item "$extractPath\miniupnpc.dll" "C:\Windows\System32\" -Force
        
        # Очистка
        Remove-Item $zipPath -Force
        Remove-Item $extractPath -Recurse -Force
        
        Write-Host "Установка завершена успешно!" -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Host "Ошибка при установке: $_" -ForegroundColor Red
        pause
        exit
    }
}

# Функция для настройки стандартных портов
function Set-DefaultPorts {
    Write-Host "`nНастройка стандартных портов..." -ForegroundColor Cyan

    $ports = @(
        @{Port="25565"; Protocol="tcp"; Description="Minecraft TCP"},
        @{Port="25565"; Protocol="udp"; Description="Minecraft UDP"},
        @{Port="7656"; Protocol="udp"; Description="ModularVoice UDP"},
        @{Port="24454"; Protocol="udp"; Description="SimpleVoice UDP"}
    )

    foreach ($p in $ports) {
        Write-Host "Настройка порта $($p.Port) $($p.Protocol)..." -ForegroundColor Yellow
        & upnpc-static -e $p.Description -d $p.Port $p.Protocol
        & upnpc-static -e $p.Description -a "`@" $p.Port $p.Port $p.Protocol
    }

    # Добавление стандартных портов в файл настроек
    $existingPorts = @()
    if (Test-Path $PORTS_FILE) {
        try {
            $existingPorts = Get-Content $PORTS_FILE | ConvertFrom-Json
        } catch {
            $existingPorts = @()
        }
    }

    # Объединить без дубликатов
    foreach ($p in $ports) {
        if (-not ($existingPorts | Where-Object {
            $_.Port -eq $p.Port -and $_.Protocol -eq $p.Protocol -and $_.Description -eq $p.Description
        })) {
            $existingPorts += $p
        }
    }

    $existingPorts | ConvertTo-Json | Set-Content $PORTS_FILE

    Write-Host "`nТекущие правила:" -ForegroundColor Green
    & upnpc-static -l
    pause
}

# Функция для работы с пользовательскими портами
function Set-CustomPorts {
    do {
        Clear-Host
        Show-Header "Настройка пользовательских портов"
        
        # Показать текущие порты
        $currentPorts = @()
        if (Test-Path $PORTS_FILE) {
            $currentPorts = Get-Content $PORTS_FILE | ConvertFrom-Json -ErrorAction SilentlyContinue
        }
        
        if ($currentPorts.Count -gt 0) {
            Write-Host "Текущие настройки портов:" -ForegroundColor Cyan
            foreach ($p in $currentPorts) {
                Write-Host ("  ├─ {0} {1} - {2}" -f $p.Port, $p.Protocol.ToUpper(), $p.Description) -ForegroundColor Gray
            }
        } else {
            Write-Host "Нет настроенных портов" -ForegroundColor Yellow
        }
        
        Write-Host "`nДоступные действия:" -ForegroundColor Cyan
        Write-Host "[ 1 ] Добавить новый порт" -ForegroundColor White
        Write-Host "[ 2 ] Удалить порт" -ForegroundColor White
        Write-Host "[ 3 ] Вернуться в главное меню" -ForegroundColor White
        
        $option = Read-Host "Выберите действие"
        
        if ($option -eq 1) {
            $port = Read-Host "Введите номер порта"
            $protocol = Read-Host "Введите протокол (tcp/udp)"
            $description = Read-Host "Введите описание"
            
            $newPort = @{Port=$port; Protocol=$protocol; Description=$description}
            $currentPorts += $newPort
            $currentPorts | ConvertTo-Json | Set-Content $PORTS_FILE
        } elseif ($option -eq 2) {
            $port = Read-Host "Введите номер порта для удаления"
            $currentPorts = $currentPorts | Where-Object { $_.Port -ne $port }
            $currentPorts | ConvertTo-Json | Set-Content $PORTS_FILE
        }
    } while ($option -ne 3)
}

# Запуск меню
do {
    Show-Menu
    $choice = Read-Host "Выберите действие"
    
    switch ($choice) {
        1 { Set-DefaultPorts }
        2 { Set-CustomPorts }
        3 { & upnpc-static -l }
        4 { & upnpc-static -L }
        5 { & upnpc-static -d 25565 udp; & upnpc-static -d 25565 tcp }
        6 { exit }
        default { Write-Host "Неверный выбор, попробуйте снова." -ForegroundColor Red }
    }
} while ($choice -ne 6)
