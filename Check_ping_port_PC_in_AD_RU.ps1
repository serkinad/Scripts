<#
.SYNOPSIS
    Проверка доступности компьютеров и портов в сети с экспортом результатов в CSV.

.DESCRIPTION
    Скрипт выполняет проверку доступности компьютеров по сети (ping) и проверяет состояние указанных портов.
    Поддерживает несколько способов получения списка компьютеров: поиск по Active Directory, по OU или ручной ввод.
    Результаты сохраняются в CSV файл с подробной статистикой.

.PARAMETER None
    Скрипт работает в интерактивном режиме без параметров.

.EXAMPLE
    .\PortChecker.ps1
    Запускает интерактивное меню для проверки доступности компьютеров и портов.

.NOTES
    Автор: Aleksandr Serkin
    Версия: 2.0
    Дата создания: 2026
    Требования: PowerShell 5.1+, модуль ActiveDirectory, права на чтение AD

    Функциональные возможности:
    1. Поддержка нескольких источников списка компьютеров
    2. Проверка пинга (всегда выполняется)
    3. Проверка одного или нескольких портов
    4. Параллельная обработка до 10 компьютеров одновременно
    5. Экспорт результатов в CSV с разделителем ";"
    6. Автоматическое создание директорий для сохранения
    7. Поддержка переменных путей ($PSScriptRoot, $home, $desktop)
    8. Подробная статистика по результатам проверки

.LINK
    https://learn.microsoft.com/powershell/
    https://learn.microsoft.com/windows-server/identity/ad-ds/get-started/ad-ds-introduction

#>

[CmdletBinding()]
param()

function Show-Menu {
    <#
    .SYNOPSIS
        Отображает главное меню скрипта.
    
    .DESCRIPTION
        Очищает консоль и выводит заголовок скрипта.
    #>
    Clear-Host
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "     ПРОВЕРКА ДОСТУПНОСТИ ПОРТОВ И ПИНГ" -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-ComputersByCriteria {
    <#
    .SYNOPSIS
        Получает список компьютеров по выбранному критерию.
    
    .DESCRIPTION
        В зависимости от выбранного критерия получает список компьютеров:
        - По части имени из Active Directory
        - По Organizational Unit (OU)
        - Ручной ввод через запятую
    
    .PARAMETER Choice
        Критерий поиска компьютеров:
        - "NamePart" - поиск по части имени
        - "OU" - поиск по OU
        - "Manual" - ручной ввод
    
    .OUTPUTS
        System.Array
        Массив строк с именами компьютеров.
    
    .EXAMPLE
        Get-ComputersByCriteria -Choice "NamePart"
        Запрашивает часть имени и возвращает список компьютеров из AD.
    #>
    param(
        [ValidateSet("NamePart", "OU", "Manual", "Back")]
        [string]$Choice
    )
    
    $computers = @()
    
    switch ($Choice) {
        "NamePart" {
            $partName = Read-Host "Введите часть имени компьютера"
            if (-not [string]::IsNullOrWhiteSpace($partName)) {
                try {
                    $computers = Get-ADComputer -LDAPFilter "(cn=*$partName*)" -ErrorAction Stop | 
                                Select-Object -ExpandProperty Name
                }
                catch {
                    Write-Host "Ошибка при поиске компьютеров: $_" -ForegroundColor Red
                    Read-Host "Нажмите Enter для продолжения"
                }
            }
        }
        
        "OU" {
            $OUPath = Read-Host "Введите путь к OU (например: ou=Computers,dc=domain,dc=local)"
            if (-not [string]::IsNullOrWhiteSpace($OUPath)) {
                try {
                    $computers = Get-ADComputer -SearchBase $OUPath -Filter * -ErrorAction Stop | 
                                Select-Object -ExpandProperty Name
                }
                catch {
                    Write-Host "Ошибка при поиске компьютеров: $_" -ForegroundColor Red
                    Read-Host "Нажмите Enter для продолжения"
                }
            }
        }
        
        "Manual" {
            $manualInput = Read-Host "Введите имена компьютеров через запятую"
            if (-not [string]::IsNullOrWhiteSpace($manualInput)) {
                $computers = $manualInput.Split(',').Trim() | Where-Object { $_ -ne '' }
            }
        }
    }
    
    return $computers
}

function Test-Port {
    <#
    .SYNOPSIS
        Проверяет доступность TCP-порта на указанном компьютере.
    
    .DESCRIPTION
        Пытается установить TCP-соединение с указанным портом на компьютере.
        Использует асинхронное соединение с таймаутом.
    
    .PARAMETER comp
        Имя компьютера или IP-адрес для проверки.
    
    .PARAMETER port
        Номер TCP-порта для проверки.
    
    .PARAMETER timeout
        Таймаут подключения в миллисекундах (по умолчанию 200 мс).
    
    .OUTPUTS
        System.Boolean
        $true - порт открыт, $false - порт закрыт или недоступен.
    
    .EXAMPLE
        Test-Port -comp "server01" -port 3389
        Проверяет доступность порта 3389 на server01.
    #>
    param(
        [string]$comp,
        [int]$port,
        [int]$timeout = 200
    )
    
    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $task = $client.ConnectAsync($comp, $port)
        $completed = $task.Wait($timeout)
        return $completed -and $client.Connected
    }
    catch { 
        return $false 
    }
    finally { 
        if ($client) { 
            $client.Dispose() 
        } 
    }
}

function Test-Ping {
    <#
    .SYNOPSIS
        Проверяет доступность компьютера с помощью ICMP ping.
    
    .DESCRIPTION
        Отправляет ICMP эхо-запрос на указанный компьютер и проверяет ответ.
    
    .PARAMETER comp
        Имя компьютера или IP-адрес для проверки.
    
    .PARAMETER count
        Количество пинг-запросов (по умолчанию 1).
    
    .PARAMETER timeout
        Таймаут ожидания ответа в миллисекундах (по умолчанию 1000 мс).
    
    .OUTPUTS
        System.Boolean
        $true - компьютер отвечает на ping, $false - нет ответа.
    
    .EXAMPLE
        Test-Ping -comp "server01"
        Проверяет доступность server01 через ping.
    #>
    param(
        [string]$comp,
        [int]$count = 1,
        [int]$timeout = 1000
    )
    
    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $reply = $ping.Send($comp, $timeout)
        return $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success
    }
    catch {
        return $false
    }
}

function Read-PathWithTabCompletion {
    <#
    .SYNOPSIS
        Запрашивает путь с поддержкой автодополнения и переменных.
    
    .DESCRIPTION
        Позволяет пользователю ввести путь с поддержкой автодополнения через Tab
        и использованием предопределенных переменных.
    
    .PARAMETER Prompt
        Текст приглашения для ввода.
    
    .PARAMETER DefaultPath
        Путь по умолчанию, который будет использоваться если ввод пустой.
    
    .OUTPUTS
        System.String
        Полный путь к файлу с расширением .csv.
    
    .EXAMPLE
        Read-PathWithTabCompletion -Prompt "Введите путь" -DefaultPath "C:\default.csv"
        Запрашивает путь у пользователя с подсказкой и значением по умолчанию.
    
    .NOTES
        Поддерживаемые переменные:
        - $PSScriptRoot - директория скрипта
        - $home - домашняя директория пользователя
        - $desktop - рабочий стол пользователя
    #>
    param(
        [string]$Prompt = "Введите путь",
        [string]$DefaultPath
    )
    
    # Используем ReadLine для поддержки автодополнения
    # Проверяем, доступен ли модуль PSReadLine
    if (Get-Module -Name PSReadLine -ErrorAction SilentlyContinue) {
        Write-Host $Prompt -ForegroundColor Yellow
        if (-not [string]::IsNullOrWhiteSpace($DefaultPath)) {
            Write-Host "По умолчанию: $DefaultPath" -ForegroundColor Gray
            Write-Host "Нажмите Enter для использования пути по умолчанию" -ForegroundColor Gray
        }
        Write-Host "Используйте Tab для автодополнения" -ForegroundColor Gray
        Write-Host "Путь: " -NoNewline -ForegroundColor White
        $userInput = [Microsoft.PowerShell.PSConsoleReadLine]::ReadLine()
        
        # Если ввод пустой, возвращаем путь по умолчанию
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return $DefaultPath
        }
        
        return $userInput
    }
    else {
        # Альтернатива без PSReadLine
        Write-Host $Prompt -ForegroundColor Yellow
        if (-not [string]::IsNullOrWhiteSpace($DefaultPath)) {
            Write-Host "По умолчанию: $DefaultPath" -ForegroundColor Gray
        }
        Write-Host "Доступные переменные: `$PSScriptRoot, `$home, `$desktop" -ForegroundColor Gray
        Write-Host "Пример: `$home\Documents\results.csv" -ForegroundColor Gray
        $userInput = Read-Host "Введите путь (или нажмите Enter для использования пути по умолчанию)"
        
        # Если ввод пустой, возвращаем путь по умолчанию
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return $DefaultPath
        }
        
        # Обрабатываем специальные переменные
        $userInput = $userInput.Replace('$PSScriptRoot', $PSScriptRoot)
        $userInput = $userInput.Replace('$home', $HOME)
        $userInput = $userInput.Replace('$desktop', [Environment]::GetFolderPath('Desktop'))
        $userInput = $userInput.Replace('`$PSScriptRoot', $PSScriptRoot)
        $userInput = $userInput.Replace('`$home', $HOME)
        $userInput = $userInput.Replace('`$desktop', [Environment]::GetFolderPath('Desktop'))
        
        # Если введен путь, но нет расширения .csv, добавляем его
        if (-not $userInput.EndsWith('.csv') -and -not [string]::IsNullOrWhiteSpace($userInput)) {
            $userInput += ".csv"
        }
        
        return $userInput
    }
}

# Главное меню
do {
    Show-Menu
    
    # Получаем список компьютеров
    $allComputers = @()
    
    # Показываем меню выбора один раз
    Show-Menu
    Write-Host "ВЫБЕРИТЕ КРИТЕРИЙ ПОИСКА КОМПЬЮТЕРОВ:" -ForegroundColor Green
    Write-Host "1. Полное или часть имени ПК" -ForegroundColor White
    Write-Host "2. По OU" -ForegroundColor White
    Write-Host "3. Имена ПК через запятую" -ForegroundColor White
    Write-Host "0. Выход" -ForegroundColor Red
    Write-Host ""
    
    $searchChoice = Read-Host "Выберите вариант (1-3 или 0)"
    
    switch ($searchChoice) {
        "1" {
            $foundComputers = Get-ComputersByCriteria -Choice "NamePart"
            if ($foundComputers.Count -gt 0) {
                $allComputers += $foundComputers
            }
        }
        
        "2" {
            $foundComputers = Get-ComputersByCriteria -Choice "OU"
            if ($foundComputers.Count -gt 0) {
                $allComputers += $foundComputers
            }
        }
        
        "3" {
            $foundComputers = Get-ComputersByCriteria -Choice "Manual"
            if ($foundComputers.Count -gt 0) {
                $allComputers += $foundComputers
            }
        }
        
        "0" {
            exit
        }
        
        default {
            Write-Host "Неверный выбор!" -ForegroundColor Red
            Read-Host "Нажмите Enter для продолжения"
            continue
        }
    }
    
    # Проверяем, что список компьютеров не пустой
    if ($allComputers.Count -eq 0) {
        Write-Host "Список компьютеров пуст! Нужно добавить хотя бы один компьютер." -ForegroundColor Red
        Write-Host ""
        
        do {
            Write-Host "Выберите действие:" -ForegroundColor Yellow
            Write-Host "1. Вернуться в главное меню" -ForegroundColor White
            Write-Host "2. Выход" -ForegroundColor White
            Write-Host ""
            
            $emptyListChoice = Read-Host "Введите 1 или 2"
            
            switch ($emptyListChoice) {
                "1" {
                    # Возвращаемся в начало главного цикла
                    $continueMain = $true
                    break
                }
                "2" {
                    exit
                }
                default {
                    Write-Host "Неверный выбор! Введите 1 или 2." -ForegroundColor Red
                    $continueMain = $false
                }
            }
        } while (-not $continueMain)
        
        if ($continueMain) {
            continue
        }
    }
    
    # Если дошли сюда, значит есть компьютеры для проверки
    Write-Host "Найдено компьютеров: $($allComputers.Count)" -ForegroundColor Green
    Read-Host "Нажмите Enter для продолжения"
    
    # Получаем порты для проверки (без портов по умолчанию)
    Show-Menu
    Write-Host "Найдено компьютеров для проверки: $($allComputers.Count)" -ForegroundColor Green
    Write-Host ""
    $portsInput = Read-Host "Введите порты для проверки через запятую (или оставьте пустым для проверки только пинга)"
    
    $ports = @()
    if (-not [string]::IsNullOrWhiteSpace($portsInput)) {
        $ports = $portsInput.Split(',').Trim() | ForEach-Object { 
            if ($_ -match '^\d+$') { [int]$_ } 
        } | Where-Object { $_ -gt 0 -and $_ -lt 65536 }
        
        if ($ports.Count -gt 0) {
            Write-Host "Будут проверены порты: $($ports -join ', ')" -ForegroundColor Yellow
        }
    }
    
    # Пинг выполняется всегда по умолчанию
    Write-Host "Пинг будет выполнен для всех компьютеров" -ForegroundColor Yellow
    
    # Запрашиваем директорию для сохранения
    Write-Host ""
    $defaultPath = Join-Path $PSScriptRoot "PortCheck_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    # Используем улучшенную функцию для ввода пути
    $savePath = Read-PathWithTabCompletion -Prompt "Введите путь для сохранения CSV файла" -DefaultPath $defaultPath
    
    # Проверяем, что путь не пустой
    if ([string]::IsNullOrWhiteSpace($savePath)) {
        Write-Host "Путь не указан. Используется путь по умолчанию." -ForegroundColor Yellow
        $savePath = $defaultPath
    }
    
    Write-Host "Файл будет сохранен как: $savePath" -ForegroundColor Green
    
    # Проверяем и создаем директорию, если нужно
    $directory = Split-Path $savePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        try {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            Write-Host "Создана директория: $directory" -ForegroundColor Green
        }
        catch {
            Write-Host "Не удалось создать директорию: $_" -ForegroundColor Red
            Write-Host "Сохранение в путь по умолчанию: $defaultPath" -ForegroundColor Yellow
            $savePath = $defaultPath
        }
    }
    
    # Выполняем проверки
    $results = @()
    $computerCount = $allComputers.Count
    $current = 0
    
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "     НАЧАЛО ПРОВЕРКИ" -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Начинаем проверку $computerCount компьютеров..." -ForegroundColor Yellow
    Write-Host "Выполняется пинг для всех компьютеров" -ForegroundColor Yellow
    if ($ports.Count -gt 0) {
        Write-Host "Проверяются порты: $($ports -join ', ')" -ForegroundColor Yellow
    }
    Write-Host ""
    
    # Определяем функции внутри ForEach-Object для видимости в параллельном контексте
    $allComputers | ForEach-Object -Parallel {
        $comp = $_
        $ports = $using:ports
        
        # Определяем функции внутри параллельного блока
        function Local-TestPing {
            param(
                [string]$comp,
                [int]$count = 1,
                [int]$timeout = 1000
            )
            
            try {
                $ping = New-Object System.Net.NetworkInformation.Ping
                $reply = $ping.Send($comp, $timeout)
                return $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success
            }
            catch {
                return $false
            }
        }
        
        function Local-TestPort {
            param(
                [string]$comp,
                [int]$port,
                [int]$timeout = 200
            )
            
            $client = $null
            try {
                $client = New-Object System.Net.Sockets.TcpClient
                $task = $client.ConnectAsync($comp, $port)
                $completed = $task.Wait($timeout)
                return $completed -and $client.Connected
            }
            catch { 
                return $false 
            }
            finally { 
                if ($client) { 
                    $client.Dispose() 
                } 
            }
        }
        
        # Создаем объект для результатов
        $result = [PSCustomObject]@{
            ComputerName = $comp
            Ping = "NotTested"
            TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Добавляем свойства для портов, если они указаны
        foreach ($port in $ports) {
            $result | Add-Member -NotePropertyName "Port_$port" -NotePropertyValue "NotTested" -Force
        }
        
        # Проверка пинга (всегда выполняется)
        $pingResult = Local-TestPing -comp $comp
        $result.Ping = if ($pingResult) { "Success" } else { "Failed" }
        
        # Проверка портов (только если указаны)
        foreach ($port in $ports) {
            $portResult = Local-TestPort -comp $comp -port $port
            $result."Port_$port" = if ($portResult) { "Open" } else { "Closed" }
        }
        
        # Возвращаем результат
        $result
        
    } -ThrottleLimit 10 | ForEach-Object {
        $results += $_
        $current++
        Write-Progress -Activity "Проверка компьютеров" -Status "Обработано: $current из $computerCount" `
                      -PercentComplete (($current / $computerCount) * 100)
    }
    
    Write-Progress -Activity "Проверка компьютеров" -Completed
    
    # Сохраняем результаты
    try {
        # Проверяем, что путь не пустой перед сохранением
        if ([string]::IsNullOrWhiteSpace($savePath)) {
            $savePath = $defaultPath
            Write-Host "Путь не указан. Используется путь по умолчанию: $savePath" -ForegroundColor Yellow
        }
        
        $results | Export-Csv -Path $savePath -Encoding UTF8 -NoTypeInformation -Delimiter ';'
        Write-Host ""
        Write-Host "=========================================" -ForegroundColor Green
        Write-Host "     РЕЗУЛЬТАТЫ ПРОВЕРКИ" -ForegroundColor Yellow
        Write-Host "=========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Результаты сохранены в: $savePath" -ForegroundColor Green
        Write-Host "Обработано компьютеров: $($results.Count)" -ForegroundColor Green
        
        # Показываем краткую статистику
        $successPing = ($results | Where-Object { $_.Ping -eq "Success" }).Count
        $failedPing = $computerCount - $successPing
        Write-Host "Успешный пинг: $successPing из $($results.Count)" -ForegroundColor Green
        Write-Host "Неуспешный пинг: $failedPing из $($results.Count)" -ForegroundColor Red
        
        foreach ($port in $ports) {
            $openPorts = ($results | Where-Object { $_."Port_$port" -eq "Open" }).Count
            $closedPorts = $computerCount - $openPorts
            Write-Host "Открыт порт $port : $openPorts из $($results.Count)" -ForegroundColor Green
            Write-Host "Закрыт порт $port : $closedPorts из $($results.Count)" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Ошибка при сохранении файла: $_" -ForegroundColor Red
        Write-Host "Попытка сохранить в путь по умолчанию..." -ForegroundColor Yellow
        try {
            $results | Export-Csv -Path $defaultPath -Encoding UTF8 -NoTypeInformation -Delimiter ';'
            Write-Host "Результаты сохранены в путь по умолчанию: $defaultPath" -ForegroundColor Green
        }
        catch {
            Write-Host "Не удалось сохранить файл даже в путь по умолчанию" -ForegroundColor Red
            Write-Host "Вывод результатов в консоль:" -ForegroundColor Yellow
            $results | Format-Table -AutoSize
        }
    }
    
    Write-Host ""
    do {
        Write-Host "Выберите действие:" -ForegroundColor Yellow
        Write-Host "1. Выполнить еще одну проверку" -ForegroundColor White
        Write-Host "2. Выход" -ForegroundColor White
        Write-Host ""
        
        $finalChoice = Read-Host "Введите 1 или 2"
        
        switch ($finalChoice) {
            "1" {
                $continueMain = $true
                break
            }
            "2" {
                Write-Host "Скрипт завершен." -ForegroundColor Green
                exit
            }
            default {
                Write-Host "Неверный выбор! Введите 1 или 2." -ForegroundColor Red
                $continueMain = $false
            }
        }
    } while (-not $continueMain)
    
} while ($continueMain)