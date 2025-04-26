# miniUPnP Manager by Seilce
# Требуются права администратора
#Requires -RunAsAdministrator
# Получаем текущую идентификацию пользователя

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
$size.Width = 100
$size.Height = 25
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

$PORTS_FILE = Join-Path $env:USERPROFILE "UPnP\nazzy_ports.txt"
$upnpcPath = Join-Path $env:USERPROFILE "UPnP\upnpc-static.exe"
$upnpfolder = Join-Path $env:USERPROFILE "UPnP\"
$upnpfoldercreate = $env:USERPROFILE
function CreateFolder {
    if (-not (Test-Path $upnpfolder)) {
        New-Item -Path $upnpfolder -ItemType Directory | Out-Null
        Write-Host "Папка 'UPnP' создана." -ForegroundColor Green
    } else {
        Write-Host "Папка 'UPnP' уже существует." -ForegroundColor Yellow
    }
}


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
    Write-Host "     ├─ 7656 TCP/UDP (ModularVoice)" -ForegroundColor Gray
    Write-Host "     └─ 24454 TCP/UDP (SimpleVoice)" -ForegroundColor Gray
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
        Copy-Item "$extractPath\upnpc-static.exe" "$upnpfolder" -Force
        Copy-Item "$extractPath\upnpc-shared.exe" "$upnpfolder" -Force
        Copy-Item "$extractPath\miniupnpc.dll" "$upnpfolder" -Force
        
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
        @{Port="24454"; Protocol="udp"; Description="SimpleVoice UDP"},
        @{Port="7656"; Protocol="tcp"; Description="ModularVoice TCP"},
        @{Port="24454"; Protocol="tcp"; Description="SimpleVoice TCP"}
    )

    foreach ($p in $ports) {
        Write-Host "Настройка порта $($p.Port) $($p.Protocol)..." -ForegroundColor Yellow
        & $upnpcPath -e $p.Description -d $p.Port $p.Protocol
        & $upnpcPath -e $p.Description -a "`@" $p.Port $p.Port $p.Protocol
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
    & $upnpcPath -l
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
        
        Write-Host "`nДоступные действия:" -ForegroundColor White
        Write-Host "[ 1 ]" -ForegroundColor Green -NoNewline
        Write-Host " Добавить порт" -ForegroundColor White
        Write-Host "[ 2 ]" -ForegroundColor Green -NoNewline
        Write-Host " Применить настройки" -ForegroundColor White
        Write-Host "[ 3 ]" -ForegroundColor Yellow -NoNewline
        Write-Host " Удалить порт" -ForegroundColor White
        Write-Host "[ 4 ]" -ForegroundColor Red -NoNewline
        Write-Host " Вернуться в главное меню" -ForegroundColor White
        Write-Host
        
        $choice = Read-Host "Выберите действие (1-4)"
        
        switch ($choice) {
            "1" { 
                Add-CustomPort -currentPorts $currentPorts
            }
            "2" { 
                if ($currentPorts.Count -gt 0) {
                    Apply-CustomPorts -ports $currentPorts
                } else {
                    Write-Host "Нет портов для применения!" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
            "3" {
                if ($currentPorts.Count -gt 0) {
                    Remove-CustomPort -currentPorts $currentPorts
                } else {
                    Write-Host "Нет портов для удаления!" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
            "4" { return }
            default {
                Write-Host "Неверный выбор!" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

# Функция для добавления пользовательского порта
function Add-CustomPort {
    param (
        [array]$currentPorts
    )
    
    do {
        $port = Read-Host "`nВведите номер порта (1-65535)"
        if ($port -match '^\d+$' -and [int]$port -ge 1 -and [int]$port -le 65535) {
            break
        }
        Write-Host "Неверный номер порта!" -ForegroundColor Red
    } while ($true)
    
    Write-Host "`nВыберите протокол:"
    Write-Host "[ 1 ]" -ForegroundColor Cyan -NoNewline
    Write-Host " TCP" -ForegroundColor White
    Write-Host "[ 2 ]" -ForegroundColor Cyan -NoNewline
    Write-Host " UDP" -ForegroundColor White
    Write-Host "[ 3 ]" -ForegroundColor Cyan -NoNewline
    Write-Host " TCP и UDP" -ForegroundColor White
    
    do {
        $protChoice = Read-Host "Ваш выбор (1-3)"
        if ($protChoice -match '^[1-3]$') {
            break
        }
        Write-Host "Неверный выбор!" -ForegroundColor Red
    } while ($true)
    
    $description = Read-Host "`nВведите описание правила"
    
    switch ($protChoice) {
        "1" { 
            $currentPorts += @{Port=$port; Protocol="tcp"; Description=$description}
        }
        "2" { 
            $currentPorts += @{Port=$port; Protocol="udp"; Description=$description}
        }
        "3" { 
            $currentPorts += @{Port=$port; Protocol="tcp"; Description="$description (TCP)"}
            $currentPorts += @{Port=$port; Protocol="udp"; Description="$description (UDP)"}
        }
    }
    
    $currentPorts | ConvertTo-Json | Set-Content $PORTS_FILE
    Write-Host "`nПорт успешно добавлен!" -ForegroundColor Green
    Start-Sleep -Seconds 1
}

# Функция для применения пользовательских портов
function Apply-CustomPorts {
    param (
        [array]$ports
    )
    
    Write-Host "`nПрименение настроек портов..." -ForegroundColor Cyan
    
    foreach ($p in $ports) {
        Write-Host "Настройка порта $($p.Port) $($p.Protocol)..." -ForegroundColor Yellow
        & $upnpcPath -e $p.Description -d $p.Port $p.Protocol
        & $upnpcPath -e $p.Description -a "`@" $p.Port $p.Port $p.Protocol
    }
    
    Write-Host "`nТекущие правила:" -ForegroundColor Green
    & $upnpcPath -l
    pause
}

# Функция для удаления пользовательского порта
function Remove-CustomPort {
    param (
        [array]$currentPorts
    )

    Clear-Host
    Show-Header "Удаление пользовательских портов"

    if ($currentPorts.Count -gt 0) {
        Write-Host "Доступные порты для удаления:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $currentPorts.Count; $i++) {
            Write-Host ("  ├─ {0}. {1} {2} - {3}" -f ($i+1), $currentPorts[$i].Port, 
                $currentPorts[$i].Protocol.ToUpper(), $currentPorts[$i].Description) -ForegroundColor Gray
        }

    } else {
        Write-Host "Нет настроенных портов" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        return
    }

    Write-Host ""
    Write-Host "`nДоступные действия:" -ForegroundColor White
    Write-Host "[ 0 ]" -ForegroundColor Red -NoNewline
    Write-Host " Отмена" -ForegroundColor White
    Write-Host ""
    $selection = Read-Host "Выберите порт для удаления (0-$($currentPorts.Count))"

    if ($selection -eq "0") { 
        return 
    } elseif ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $currentPorts.Count) {
        $index = [int]$selection - 1
        $port = $currentPorts[$index]

        Write-Host "`nВы уверены, что хотите удалить порт $($port.Port) $($port.Protocol)?" -ForegroundColor Yellow
        $confirmation = Read-Host "Подтвердите (y/n)"

        if ($confirmation -eq 'y') {
            # Удаление правила из UPnP
            Write-Host "Удаление правила из UPnP: $($port.Port) $($port.Protocol)" -ForegroundColor DarkGray
            & $upnpcPath -e $port.Description -d $port.Port $port.Protocol

            # Удаление порта из массива
            $newPorts = @($currentPorts | Where-Object { 
                -not (($_.Port -eq $port.Port) -and ($_.Protocol -eq $port.Protocol) -and ($_.Description -eq $port.Description))
            })

            # Сохраняем обновленные порты
            if ($newPorts.Count -gt 0) {
                $newPorts | ConvertTo-Json | Set-Content $PORTS_FILE
            } else {
                if (Test-Path $PORTS_FILE) { Remove-Item $PORTS_FILE -Force }
            }

            Write-Host "`nПорт и правило успешно удалены!" -ForegroundColor Green
            Start-Sleep -Seconds 1
        }
    } else {
        Write-Host "Неверный выбор!" -ForegroundColor Red
        Start-Sleep -Seconds 1
    }
}

# Обновляем функцию удаления всех правил
function Remove-AllRules {
    Write-Host "`nУдаление всех правил..." -ForegroundColor Yellow
    
    # Удаление стандартных портов
    $defaultPorts = @(
        @{Port="25565"; Protocol="tcp"; Desc="Minecraft TCP"},
        @{Port="25565"; Protocol="udp"; Desc="Minecraft UDP"},
        @{Port="7656"; Protocol="udp"; Desc="ModularVoice UDP"},
        @{Port="24454"; Protocol="udp"; Desc="SimpleVoice UDP"}
    )
    
    foreach ($p in $defaultPorts) {
        & $upnpcPath -e $p.Desc -d $p.Port $p.Protocol
    }
    
    # Удаление пользовательских портов
    if (Test-Path $PORTS_FILE) {
        $customPorts = Get-Content $PORTS_FILE | ConvertFrom-Json -ErrorAction SilentlyContinue
        foreach ($p in $customPorts) {
            & $upnpcPath -e $p.Description -d $p.Port $p.Protocol
        }
        Remove-Item $PORTS_FILE -Force
    }
    
    Write-Host "Все правила были удалены!" -ForegroundColor Green
    pause
}

# Основной цикл программы
function Start-MainLoop {
    Test-Upnpc
    
    do {
        Show-Menu
        $choice = Read-Host "Выберите действие (1-6)"
        
        switch ($choice) {
            "1" { Set-DefaultPorts }
            "2" { Set-CustomPorts }
            "3" { 
                Write-Host "`nПроверка статуса подключения..." -ForegroundColor Yellow
                & $upnpcPath -s
                pause
            }
            "4" {
                Write-Host "`nТекущие правила:" -ForegroundColor Green
                & $upnpcPath -l
                pause
            }
            "5" { Remove-AllRules }
            "6" { return }
            default { 
                Write-Host "Неверный выбор!" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

# Запуск программы
Start-MainLoop 
