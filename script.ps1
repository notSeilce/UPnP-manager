#Requires -RunAsAdministrator

# miniUPnP Manager by Seilce
# Требуются права администратора

# Установка кодировки для корректного отображения русского языка
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# Установка стандартной темы консоли для PowerShell (Этот блок может быть избыточен, если реестр уже настроен)

# --- Начало настройки вида консоли ---
try {
    # Переход в нужный раздел реестра
    Set-Location HKCU:\Console

    # Создание раздела профиля для powershell.exe, если он не существует
    $psProfileKey = '.\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe'
    if (-not (Test-Path $psProfileKey)) {
        New-Item $psProfileKey -Force | Out-Null
    }

    Set-Location $psProfileKey

    # Установка размера окна (может не сработать во всех терминалах, например, Windows Terminal)
    # $size = $host.UI.RawUI.WindowSize
    # $size.Width = 120
    # $size.Height = 30
    # $host.UI.RawUI.WindowSize = $size

    # Установка шрифта Consolas 16 (требует прав админа для HKCU:\Console)
    New-ItemProperty -Path . -Name FaceName   -Value "Consolas"     -PropertyType String -Force -ErrorAction SilentlyContinue
    New-ItemProperty -Path . -Name FontFamily -Value 0x00000036     -PropertyType DWord  -Force -ErrorAction SilentlyContinue # 54 (FF_MODERN)
    New-ItemProperty -Path . -Name FontSize   -Value 0x00100000     -PropertyType DWord  -Force -ErrorAction SilentlyContinue # 0x0010 = 16pt (16 pixels high)
    New-ItemProperty -Path . -Name FontWeight -Value 0x00000190     -PropertyType DWord  -Force -ErrorAction SilentlyContinue # 400 = normal

    # Попытка применить шрифт немедленно (может не работать или требовать перезапуска)
    # Set-ConsoleFont 16 # Эта команда не является стандартной PowerShell

    # Возвращаемся обратно
    Set-Location $env:USERPROFILE
} catch {
    Write-Warning "Не удалось изменить настройки реестра для консоли. Ошибка: $($_.Exception.Message)"
    Write-Warning "Продолжаем работу со стандартными настройками."
    Set-Location $env:USERPROFILE # Убедимся, что вернулись в домашнюю директорию
}
# --- Конец настройки вида консоли ---

# Устанавливаем визуальные параметры консоли для текущей сессии
(Get-Host).UI.RawUI.ForegroundColor = "White"
(Get-Host).UI.RawUI.BackgroundColor = "Black"
(Get-Host).UI.RawUI.CursorSize = 10
(Get-Host).UI.RawUI.WindowTitle = "miniUPnP Manager by Seilce"
Clear-Host

# Путь к файлу с сохраненными портами
$PORTS_FILE = "C:\Program Files\nazzy_ports.txt" # Consider placing this in $env:APPDATA or $env:PROGRAMDATA for better practices

# Функция для создания красивого заголовка
function Show-Header {
    param (
        [string]$Title
    )
    $width = $host.UI.RawUI.WindowSize.Width
    Write-Host "`n"
    Write-Host ($Title) -ForegroundColor Yellow
    Write-Host ("-" * $Title.Length) -ForegroundColor Yellow # Добавим подчеркивание
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
    # Ищем upnpc-static.exe в системных путях
    $upnpcPath = Get-Command upnpc-static.exe -ErrorAction SilentlyContinue
    if (-not $upnpcPath) {
        Write-Host "upnpc не установлен или не найден в PATH. Начинаем установку..." -ForegroundColor Yellow
        Install-Upnpc
    } else {
        Write-Host "upnpc найден: $($upnpcPath.Source)" -ForegroundColor Green
    }
}

# Функция установки upnpc
function Install-Upnpc {
    try {
        Write-Host "Скачивание upnpc..." -ForegroundColor Cyan
        # Используем более надежный источник, если официальный недоступен или устарел
        # Ссылка на 2.2.4 на момент проверки (2023):
        $url = "https://miniupnp.tuxfamily.org/files/download.php?file=upnpc-exe-win32-2.2.4.zip"
        # Старая ссылка (на всякий случай): $url = "http://miniupnp.free.fr/files/upnpc-exe-win32-20220515.zip"
        $zipPath = Join-Path $env:TEMP "upnpc.zip"
        $extractPath = Join-Path $env:TEMP "temp_upnpc"

        # Скачивание файла
        Write-Host "Загрузка с $url ..." -ForegroundColor DarkCyan
        Invoke-WebRequest -Uri $url -OutFile $zipPath

        # Создание временной директории и распаковка
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

        # Ищем .exe и .dll файлы в распакованной директории (имена могут меняться)
        $exeFiles = Get-ChildItem -Path $extractPath -Filter *.exe -Recurse
        $dllFiles = Get-ChildItem -Path $extractPath -Filter *.dll -Recurse

        if ($exeFiles.Count -eq 0) {
            throw "Не найдены исполняемые файлы (*.exe) в скачанном архиве."
        }
        if ($dllFiles.Count -eq 0) {
            throw "Не найдены библиотеки (*.dll) в скачанном архиве."
        }

        # Копирование файлов в System32 (требует прав админа)
        Write-Host "Копирование файлов в C:\Windows\System32..." -ForegroundColor Cyan
        foreach ($file in $exeFiles) {
            Write-Host "  Копирую $($file.Name)" -ForegroundColor Gray
            Copy-Item $file.FullName "C:\Windows\System32\" -Force
        }
        foreach ($file in $dllFiles) {
             Write-Host "  Копирую $($file.Name)" -ForegroundColor Gray
            Copy-Item $file.FullName "C:\Windows\System32\" -Force
        }

        # Очистка
        Remove-Item $zipPath -Force
        Remove-Item $extractPath -Recurse -Force

        Write-Host "Установка завершена успешно!" -ForegroundColor Green
        Write-Host "Пожалуйста, перезапустите скрипт, если возникнут проблемы с поиском upnpc." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
    }
    catch {
        Write-Host "Ошибка при установке: $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.InnerException) {
             Write-Host "Внутренняя ошибка: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
        pause
        exit
    }
}

# === ИЗМЕНЕНИЕ 1: Функция для чтения JSON с -Raw и обработкой ошибок ===
function Get-PortsFromFile {
    param(
        [string]$FilePath = $PORTS_FILE
    )
    $ports = @() # Инициализируем пустым массивом
    if (Test-Path $FilePath) {
        try {
            # Читаем весь файл как одну строку
            $jsonText = Get-Content -Path $FilePath -Raw -Encoding UTF8
            # Проверяем, не пустой ли файл или содержит только пробелы
            if ($jsonText -match '\S') { # \S - любой не пробельный символ
                # Пытаемся конвертировать из JSON
                $ports = $jsonText | ConvertFrom-Json -ErrorAction Stop
                # Важно: Если в файле был сохранен только один объект, ConvertFrom-Json вернет PSCustomObject, а не массив.
                # Приводим к массиву для консистентности.
                if ($ports -isnot [array]) {
                    $ports = @($ports)
                }
            } else {
                 # Файл существует, но пуст
                 Write-Host "Файл настроек $FilePath пуст." -ForegroundColor Yellow
            }
        }
        catch {
            # Ошибка при чтении или парсинге JSON
            Write-Warning "Не удалось прочитать или невалидный JSON в файле $FilePath. Сбрасываем список портов. Ошибка: $($_.Exception.Message)"
            # Оставляем $ports пустым массивом
            # Можно добавить логирование или бэкап файла при желании
            # Rename-Item -Path $FilePath -NewName "$FilePath.bak" -ErrorAction SilentlyContinue
        }
    }
    # else { # Файл не существует, $ports уже @() }

    return $ports
}

# Функция для настройки стандартных портов
function Set-DefaultPorts {
    Write-Host "`nНастройка стандартных портов..." -ForegroundColor Cyan

    $defaultPortsToAdd = @(
        @{Port="25565"; Protocol="tcp"; Description="Minecraft TCP"},
        @{Port="25565"; Protocol="udp"; Description="Minecraft UDP"},
        @{Port="7656"; Protocol="udp"; Description="ModularVoice UDP"},
        @{Port="24454"; Protocol="udp"; Description="SimpleVoice UDP"},
        @{Port="7656"; Protocol="tcp"; Description="ModularVoice TCP"},
        @{Port="24454"; Protocol="tcp"; Description="SimpleVoice TCP"}
    )

    # Получаем текущие порты из файла, используя новую функцию
    $existingPorts = Get-PortsFromFile

    $portsApplied = $false
    $portsAddedToFile = $false

    # Применяем правила через UPnP и добавляем в список для сохранения
    foreach ($p in $defaultPortsToAdd) {
        Write-Host "Настройка правила для порта $($p.Port) $($p.Protocol)..." -ForegroundColor Yellow
        # Сначала пытаемся удалить старое правило (на случай если описание изменилось или правило "зависло")
        & upnpc-static -e $p.Description -d $p.Port $p.Protocol # | Out-Null # Скрыть вывод upnpc?
        # Добавляем новое правило
        & upnpc-static -e $p.Description -a ([System.Net.Dns]::GetHostByName($env:computerName)).AddressList[0].IPAddressToString $p.Port $p.Port $p.Protocol 0 # | Out-Null # Указываем локальный IP явно
        $portsApplied = $true

        # Проверяем, существует ли уже такой порт в сохраненных настройках
        $portExistsInFile = $existingPorts | Where-Object {
            $_.Port -eq $p.Port -and $_.Protocol -eq $p.Protocol # Проверяем только порт и протокол, описание может быть любым
        }

        if (-not $portExistsInFile) {
            $existingPorts += $p
            $portsAddedToFile = $true
        } elseif (($existingPorts | Where-Object { $_.Port -eq $p.Port -and $_.Protocol -eq $p.Protocol }).Description -ne $p.Description) {
            # Если порт есть, но описание отличается - обновляем описание в файле
             Write-Host "  Обновление описания для $($p.Port) $($p.Protocol) в файле настроек." -ForegroundColor DarkCyan
             $existingPorts = $existingPorts | ForEach-Object {
                 if ($_.Port -eq $p.Port -and $_.Protocol -eq $p.Protocol) {
                     $_.Description = $p.Description
                 }
                 $_ # Выводим объект дальше по конвейеру
             }
             $portsAddedToFile = $true
        }
    }

    # === ИЗМЕНЕНИЕ 3: Сохранение JSON с -Depth и -Encoding UTF8 ===
    if ($portsAddedToFile) {
        Write-Host "`nСохранение обновленного списка портов в $PORTS_FILE..." -ForegroundColor DarkCyan
        try {
            $existingPorts | ConvertTo-Json -Depth 5 | Set-Content -Path $PORTS_FILE -Encoding UTF8 -Force
            Write-Host "Список портов сохранен." -ForegroundColor Green
        } catch {
             Write-Warning "Не удалось сохранить список портов в $PORTS_FILE. Ошибка: $($_.Exception.Message)"
        }
    } else {
         Write-Host "`nСтандартные порты уже присутствуют в файле настроек." -ForegroundColor Gray
    }

    if ($portsApplied) {
        Write-Host "`nСтандартные правила UPnP применены." -ForegroundColor Green
        Write-Host "`nТекущие правила на роутере:" -ForegroundColor Cyan
        & upnpc-static -l
    } else {
         Write-Host "`nНе удалось применить правила UPnP." -ForegroundColor Red
    }
    pause
}

# Функция для работы с пользовательскими портами
function Set-CustomPorts {
    do {
        Clear-Host
        Show-Header "Настройка пользовательских портов"

        # Получаем текущие порты из файла
        $currentPorts = Get-PortsFromFile

        if ($currentPorts.Count -gt 0) {
            Write-Host "Текущие настройки портов из файла $PORTS_FILE:" -ForegroundColor Cyan
            # Сортируем для удобства
            $currentPorts | Sort-Object @{Expression={[int]$_.Port}}, Protocol | ForEach-Object {
                Write-Host ("  ├─ {0} {1} - {2}" -f $_.Port, $_.Protocol.ToUpper(), $_.Description) -ForegroundColor Gray
            }
        } else {
            Write-Host "Файл настроек $PORTS_FILE пуст или не найден." -ForegroundColor Yellow
        }

        Write-Host "`nДоступные действия:" -ForegroundColor White
        Write-Host "[ 1 ]" -ForegroundColor Green -NoNewline
        Write-Host " Добавить порт в файл" -ForegroundColor White
        Write-Host "[ 2 ]" -ForegroundColor Green -NoNewline
        Write-Host " Применить ВСЕ порты из файла через UPnP" -ForegroundColor White
        Write-Host "[ 3 ]" -ForegroundColor Yellow -NoNewline
        Write-Host " Удалить порт из файла (и попытаться удалить правило UPnP)" -ForegroundColor White
        Write-Host "[ 4 ]" -ForegroundColor Red -NoNewline
        Write-Host " Вернуться в главное меню" -ForegroundColor White
        Write-Host

        $choice = Read-Host "Выберите действие (1-4)"

        switch ($choice) {
            "1" {
                # Add-CustomPort теперь возвращает обновленный массив
                $currentPorts = Add-CustomPort -portsArray $currentPorts
            }
            "2" {
                if ($currentPorts.Count -gt 0) {
                    Apply-CustomPorts -ports $currentPorts
                } else {
                    Write-Host "Файл настроек пуст. Нечего применять!" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
            "3" {
                if ($currentPorts.Count -gt 0) {
                    # Remove-CustomPort теперь возвращает обновленный массив
                    $currentPorts = Remove-CustomPort -portsArray $currentPorts
                } else {
                    Write-Host "Файл настроек пуст. Нечего удалять!" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
            "4" { return } # Выход из функции Set-CustomPorts
            default {
                Write-Host "Неверный выбор!" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
        # Цикл продолжается, пока пользователь не выберет '4'
    } while ($true)
}

# Функция для добавления пользовательского порта
function Add-CustomPort {
    param (
        [Parameter(Mandatory=$true)]
        [array]$portsArray # Принимаем массив
    )

    $port = ''
    do {
        $portInput = Read-Host "`nВведите номер порта (1-65535) или 'q' для отмены"
        if ($portInput -eq 'q') { return $portsArray } # Возвращаем исходный массив при отмене
        if ($portInput -match '^\d+$' -and [int]$portInput -ge 1 -and [int]$portInput -le 65535) {
            $port = [int]$portInput
            break
        }
        Write-Host "Неверный номер порта! Введите число от 1 до 65535." -ForegroundColor Red
    } while ($true)

    Write-Host "`nВыберите протокол:"
    Write-Host "[ 1 ]" -ForegroundColor Cyan -NoNewline
    Write-Host " TCP" -ForegroundColor White
    Write-Host "[ 2 ]" -ForegroundColor Cyan -NoNewline
    Write-Host " UDP" -ForegroundColor White
    Write-Host "[ 3 ]" -ForegroundColor Cyan -NoNewline
    Write-Host " TCP и UDP" -ForegroundColor White
    Write-Host "[ q ]" -ForegroundColor Yellow -NoNewline
    Write-Host " Отмена" -ForegroundColor White

    $protChoice = ''
    do {
        $protChoice = Read-Host "Ваш выбор (1-3 или q)"
        if ($protChoice -eq 'q') { return $portsArray } # Отмена
        if ($protChoice -match '^[1-3]$') {
            break
        }
        Write-Host "Неверный выбор!" -ForegroundColor Red
    } while ($true)

    # === ИЗМЕНЕНИЕ 2: Принудительное приведение описания к строке ===
    $rawDesc     = Read-Host "`nВведите описание правила (может быть любым текстом, включая цифры)"
    if ($rawDesc -eq 'q') { return $portsArray } # Отмена

    # Убеждаемся, что описание всегда строка, даже если это только цифры
    $description = [string]$rawDesc
    # Дополнительно, можно обрезать пробелы по краям
    $description = $description.Trim()
    if ([string]::IsNullOrWhiteSpace($description)) {
        $description = "Port $port Custom Rule" # Ставим описание по умолчанию, если введено пустое
        Write-Host "Описание не указано, установлено значение по умолчанию: '$description'" -ForegroundColor Yellow
    }


    $added = $false
    switch ($protChoice) {
        "1" {
            # Проверяем дубликат перед добавлением
            if (-not ($portsArray | Where-Object {$_.Port -eq $port -and $_.Protocol -eq "tcp"})) {
                 $portsArray += @{Port=$port; Protocol="tcp"; Description=$description}
                 $added = $true
            } else { Write-Host "Порт $port TCP уже существует в списке." -ForegroundColor Yellow }
        }
        "2" {
             if (-not ($portsArray | Where-Object {$_.Port -eq $port -and $_.Protocol -eq "udp"})) {
                $portsArray += @{Port=$port; Protocol="udp"; Description=$description}
                 $added = $true
            } else { Write-Host "Порт $port UDP уже существует в списке." -ForegroundColor Yellow }
        }
        "3" {
            $addedTCP = $false
            $addedUDP = $false
            if (-not ($portsArray | Where-Object {$_.Port -eq $port -and $_.Protocol -eq "tcp"})) {
                # Явно приводим к строке и здесь
                $portsArray += @{Port=$port; Protocol="tcp"; Description=[string]"$description (TCP)"}
                $addedTCP = $true
            } else { Write-Host "Порт $port TCP уже существует в списке." -ForegroundColor Yellow }

            if (-not ($portsArray | Where-Object {$_.Port -eq $port -and $_.Protocol -eq "udp"})) {
                $portsArray += @{Port=$port; Protocol="udp"; Description=[string]"$description (UDP)"}
                $addedUDP = $true
            } else { Write-Host "Порт $port UDP уже существует в списке." -ForegroundColor Yellow }
            $added = $addedTCP -or $addedUDP
        }
    }

    if ($added) {
        # === ИЗМЕНЕНИЕ 3 (повторно): Сохранение JSON ===
        try {
            $portsArray | Sort-Object @{Expression={[int]$_.Port}}, Protocol | ConvertTo-Json -Depth 5 | Set-Content -Path $PORTS_FILE -Encoding UTF8 -Force
            Write-Host "`nПорт(ы) успешно добавлен(ы) в файл $PORTS_FILE!" -ForegroundColor Green
        } catch {
            Write-Warning "Не удалось сохранить обновленный список портов в $PORTS_FILE. Ошибка: $($_.Exception.Message)"
            # При ошибке сохранения, возможно, стоит вернуть исходный массив? Зависит от требований.
            # return $args[0] # Вернуть исходный массив (переданный как $portsArray)
        }
        Start-Sleep -Seconds 1
    } else {
         Write-Host "`nУказанный порт(ы) уже существует в файле. Изменения не сохранены." -ForegroundColor Yellow
         Start-Sleep -Seconds 2
    }

    # Возвращаем (возможно, обновленный) массив
    return $portsArray
}


# Функция для применения пользовательских портов (из файла) через UPnP
function Apply-CustomPorts {
    param (
        [Parameter(Mandatory=$true)]
        [array]$ports
    )

    Write-Host "`nПрименение настроек портов из файла $PORTS_FILE через UPnP..." -ForegroundColor Cyan

    if ($ports.Count -eq 0) {
        Write-Host "Список портов пуст. Нечего применять." -ForegroundColor Yellow
        pause
        return
    }

    # Получаем локальный IP адрес для команды -a
    $localIp = ''
    try {
        # Пытаемся получить IPv4 адрес
        $localIp = (Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias (Get-NetIPConfiguration | Where-Object {$_.IPv4DefaultGateway -ne $null}).InterfaceAlias | Select-Object -First 1).IPAddress
        if (-not $localIp) {
            # Если не вышло, пробуем старый метод (может вернуть IPv6)
            $localIp = ([System.Net.Dns]::GetHostByName($env:computerName)).AddressList[0].IPAddressToString
        }
         Write-Host "Обнаружен локальный IP: $localIp" -ForegroundColor DarkGray
    } catch {
        Write-Warning "Не удалось автоматически определить локальный IP адрес. Ошибка: $($_.Exception.Message)"
        $localIp = Read-Host "Пожалуйста, введите ваш локальный IP адрес вручную (например, 192.168.1.100)"
        if (-not ($localIp -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$')) {
             Write-Error "Введен неверный формат IP адреса. Применение правил отменено."
             pause
             return
        }
    }


    foreach ($p in $ports) {
        # Проверяем, что объект порта содержит все нужные свойства
        if ($p -and $p.PSObject.Properties['Port'] -and $p.PSObject.Properties['Protocol'] -and $p.PSObject.Properties['Description']) {
             Write-Host "Настройка правила для порта $($p.Port) $($p.Protocol) '$($p.Description)'..." -ForegroundColor Yellow
             # Сначала удаляем старое правило
             & upnpc-static -e $p.Description -d $p.Port $p.Protocol # | Out-Null
             # Добавляем новое правило с указанием локального IP
             & upnpc-static -e $p.Description -a $localIp $p.Port $p.Port $p.Protocol 0 # | Out-Null
        } else {
             Write-Warning "Пропущена некорректная запись в массиве портов."
        }
    }

    Write-Host "`nПрименение правил UPnP завершено." -ForegroundColor Green
    Write-Host "`nТекущие правила на роутере (может потребоваться время для обновления):" -ForegroundColor Cyan
    & upnpc-static -l
    pause
}

# Функция для удаления пользовательского порта
function Remove-CustomPort {
    param (
        [Parameter(Mandatory=$true)]
        [array]$portsArray # Принимаем массив
    )

    if ($portsArray.Count -eq 0) {
        Write-Host "Список портов пуст. Нечего удалять." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        return $portsArray # Возвращаем пустой массив
    }

    Clear-Host
    Show-Header "Удаление пользовательского порта из файла"

    Write-Host "Доступные порты для удаления (из файла $PORTS_FILE):" -ForegroundColor Cyan
    # Сортируем и нумеруем
    $sortedPorts = $portsArray | Sort-Object @{Expression={[int]$_.Port}}, Protocol
    for ($i = 0; $i -lt $sortedPorts.Count; $i++) {
        Write-Host ("  [{0}] {1} {2} - {3}" -f ($i+1), $sortedPorts[$i].Port,
            $sortedPorts[$i].Protocol.ToUpper(), $sortedPorts[$i].Description) -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "[ 0 ]" -ForegroundColor Red -NoNewline
    Write-Host " Отмена" -ForegroundColor White
    Write-Host ""

    $selection = ''
    do {
        $selection = Read-Host "Выберите номер порта для удаления (1-$($sortedPorts.Count)) или 0 для отмены"
        if ($selection -eq "0") {
            return $portsArray # Возвращаем исходный массив при отмене
        }
        if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $sortedPorts.Count) {
            break
        }
         Write-Host "Неверный выбор!" -ForegroundColor Red
    } while ($true)


    $index = [int]$selection - 1
    $portToRemove = $sortedPorts[$index] # Получаем объект порта для удаления из отсортированного списка

    Write-Host "`nВы уверены, что хотите удалить порт $($portToRemove.Port) $($portToRemove.Protocol) ('$($portToRemove.Description)') из файла?" -ForegroundColor Yellow
    Write-Host "Это также попытается удалить соответствующее правило UPnP с роутера." -ForegroundColor Yellow
    $confirmation = Read-Host "Подтвердите (y/n)"

    if ($confirmation -eq 'y') {
        # Удаление правила из UPnP
        Write-Host "Попытка удаления правила UPnP: $($portToRemove.Port) $($portToRemove.Protocol) '$($portToRemove.Description)'..." -ForegroundColor DarkGray
        & upnpc-static -e $portToRemove.Description -d $portToRemove.Port $portToRemove.Protocol # | Out-Null

        # Удаление порта из исходного массива $portsArray (не из $sortedPorts!)
        # Используем сравнение по всем ключевым полям для надежности, или по ссылке если уверены
        $newPorts = @($portsArray | Where-Object {
             -not ($_.Port -eq $portToRemove.Port -and $_.Protocol -eq $portToRemove.Protocol -and $_.Description -eq $portToRemove.Description)
           # Или если уверены что объект уникален: $_ -ne $portToRemove
        })

        # === ИЗМЕНЕНИЕ 3 (повторно): Сохранение JSON ===
        try {
            if ($newPorts.Count -gt 0) {
                $newPorts | Sort-Object @{Expression={[int]$_.Port}}, Protocol | ConvertTo-Json -Depth 5 | Set-Content -Path $PORTS_FILE -Encoding UTF8 -Force
            } else {
                # Если удалили последний порт, очищаем файл
                if (Test-Path $PORTS_FILE) {
                    Write-Host "Удален последний порт, очищаем файл $PORTS_FILE." -ForegroundColor Yellow
                    Clear-Content -Path $PORTS_FILE -Force
                    # Или можно удалить файл: Remove-Item $PORTS_FILE -Force
                }
            }
             Write-Host "`nПорт успешно удален из файла $PORTS_FILE!" -ForegroundColor Green
        } catch {
             Write-Warning "Не удалось сохранить изменения в файле $PORTS_FILE после удаления порта. Ошибка: $($_.Exception.Message)"
             # Возвращаем исходный массив в случае ошибки сохранения
             return $portsArray
        }

        Start-Sleep -Seconds 1
        # Возвращаем обновленный массив
        return $newPorts
    } else {
        Write-Host "Удаление отменено." -ForegroundColor Gray
        Start-Sleep -Seconds 1
        # Возвращаем исходный массив, так как ничего не изменилось
        return $portsArray
    }
}

# Обновленная функция удаления всех правил
function Remove-AllRules {
    Write-Host "`nУдаление ВСЕХ правил UPnP, найденных в файле $PORTS_FILE..." -ForegroundColor Yellow
    Write-Host "Это НЕ удаляет правила, добавленные другими приложениями." -ForegroundColor Yellow

    # Получаем список портов из файла
    $portsToRemove = Get-PortsFromFile

    if ($portsToRemove.Count -eq 0) {
         Write-Host "Файл настроек $PORTS_FILE пуст или не найден. Нет правил для удаления по этому списку." -ForegroundColor Yellow
         pause
         return
    }

    $confirmation = Read-Host "Вы уверены, что хотите попытаться удалить все $($portsToRemove.Count) правил из файла '$PORTS_FILE' с роутера? (y/n)"
    if ($confirmation -ne 'y') {
        Write-Host "Удаление отменено." -ForegroundColor Gray
        pause
        return
    }

    $removedCount = 0
    foreach ($p in $portsToRemove) {
        if ($p -and $p.PSObject.Properties['Port'] -and $p.PSObject.Properties['Protocol'] -and $p.PSObject.Properties['Description']) {
             Write-Host "Удаление правила: $($p.Port) $($p.Protocol) '$($p.Description)'..." -ForegroundColor DarkGray
            & upnpc-static -e $p.Description -d $p.Port $p.Protocol # | Out-Null
            $removedCount++
        } else {
             Write-Warning "Пропущена некорректная запись при попытке удаления."
        }
    }

    if ($removedCount -gt 0) {
         Write-Host "`nПопытка удаления $removedCount правил UPnP завершена." -ForegroundColor Green
         # Очищаем файл настроек после успешного удаления правил
         $clearFile = Read-Host "Очистить файл настроек '$PORTS_FILE' после удаления правил? (y/n)"
         if ($clearFile -eq 'y') {
             try {
                 Clear-Content -Path $PORTS_FILE -Force
                 Write-Host "Файл $PORTS_FILE очищен." -ForegroundColor Green
             } catch {
                 Write-Warning "Не удалось очистить файл $PORTS_FILE. Ошибка: $($_.Exception.Message)"
             }
         }
    } else {
         Write-Host "Не было найдено корректных записей для удаления правил." -ForegroundColor Yellow
    }

    Write-Host "`nТекущие правила на роутере:" -ForegroundColor Cyan
    & upnpc-static -l
    pause
}

# Основной цикл программы
function Start-MainLoop {
    Test-Upnpc # Проверяем наличие upnpc при запуске

    do {
        Show-Menu
        $choice = Read-Host "Выберите действие (1-6)"

        switch ($choice) {
            "1" { Set-DefaultPorts }
            "2" { Set-CustomPorts }
            "3" {
                Write-Host "`nПроверка статуса подключения UPnP..." -ForegroundColor Yellow
                & upnpc-static -s
                pause
            }
            "4" {
                Write-Host "`nТекущие правила UPnP на роутере:" -ForegroundColor Green
                & upnpc-static -l
                pause
            }
            "5" { Remove-AllRules }
            "6" {
                Write-Host "`nВыход из программы." -ForegroundColor Cyan
                return # Выход из функции Start-MainLoop и завершение скрипта
            }
            default {
                Write-Host "Неверный выбор! Пожалуйста, введите число от 1 до 6." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($true) # Бесконечный цикл до выбора "6"
}

# --- Запуск программы ---
Start-MainLoop
