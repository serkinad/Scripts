<#
.SYNOPSIS
    Скрипт для автоматизированного развертывания тестовых стендов в Hyper-V.

.DESCRIPTION
    Скрипт создает виртуальные машины для тестовых стендов с гибкой конфигурацией.
    Пользователь задает имя стенда и количество серверов каждого типа.
    Все виртуальные машины создаются на основе родительского differencing-диска.

.PARAMETER Нет параметров
    Скрипт работает в интерактивном режиме, запрашивая данные у пользователя.

.EXAMPLE
    PS C:\> .\Create-TestStand.ps1
    
    === Создание тестового стенда ===
    Введите 'X' на любом этапе для выхода из скрипта
    
    === ШАГ 1: Имя стенда ===
    Введите X для выхода из скрипта
    Введите имя стенда (например: mail): mail
    
    === ШАГ 2: Контроллеры домена ===
    Доступные опции:
      1 - создать 1 сервер(ов)
      2 - создать 2 сервер(ов)
      X - ВЫХОД из скрипта
    Сколько нужно контроллеров домена?: 1
    ...

.NOTES
    Автор: Aleksandr Serkin
    Версия: 1.0
    Дата создания: 14.01.2026
    
    ТРЕБОВАНИЯ:
    1. Hyper-V должен быть установлен и включен
    2. PowerShell запущен от имени администратора
    3. Родительский VHD должен существовать по пути: C:\vm\parent\parent.vhdx
    4. Достаточно свободного места на диске
    
    ПРИМЕЧАНИЯ:
    - Все виртуальные машины создаются как Generation 2
    - Используются differencing-диски для экономии места
    - На каждом этапе можно выйти из скрипта, введя 'X'
    
.LINK
    [Документация Hyper-V]: https://docs.microsoft.com/ru-ru/windows-server/virtualization/hyper-v/hyper-v-technology-overview
#>

# Скрипт для автоматизации разворачивания тестовых стендов

# Функция для запроса числового ввода с возможностью выхода
function Get-NumberInput {
    param(
        [string]$Prompt,
        [int]$Min,
        [int]$Max
    )
    
    Write-Host "`nДоступные опции:" -ForegroundColor Yellow
    for ($i = $Min; $i -le $Max; $i++) {
        Write-Host "  $i" -ForegroundColor White -NoNewline
        Write-Host " - создать $i сервер(ов)" -ForegroundColor Gray
    }
    Write-Host "  X" -ForegroundColor White -NoNewline
    Write-Host " - ВЫХОД из скрипта" -ForegroundColor Red
    
    do {
        $input = Read-Host $Prompt
        
        # Проверка на выход
        if ($input -eq "X" -or $input -eq "x") {
            Write-Host "Завершение работы скрипта..." -ForegroundColor Red
            exit
        }
        
        $number = $input -as [int]
        
        if ($number -eq $null -or $number -lt $Min -or $number -gt $Max) {
            Write-Host "Ошибка: введите число от $Min до $Max или X для выхода из скрипта" -ForegroundColor Red
        }
    } while ($number -eq $null -or $number -lt $Min -or $number -gt $Max)
    
    return $number
}

# Функция для создания виртуальной машины
function Create-VirtualMachine {
    <#
    .SYNOPSIS
        Создает виртуальную машину с заданными параметрами.
    
    .DESCRIPTION
        Создает differencing-диск на основе родительского VHD и настраивает виртуальную машину.
    
    .PARAMETER vmName
        Имя виртуальной машины.
    
    .PARAMETER vmType
        Тип виртуальной машины (используется в имени VHD файла).
    
    .PARAMETER cpuCount
        Количество виртуальных процессоров.
    
    .PARAMETER memoryGB
        Объем оперативной памяти в гигабайтах.
    
    .EXAMPLE
        Create-VirtualMachine -vmName "mail.dc1" -vmType "dc1" -cpuCount 2 -memoryGB 8
    #>
    
    param(
        [string]$vmName,
        [string]$vmType,
        [int]$cpuCount,
        [int]$memoryGB
    )
    
    $vmPath = "c:\vm\$vmName"
    $vhdPath = "$vmPath\vhd\$vmType.vhdx"
    
    Write-Host "Создание $vmName..." -ForegroundColor Green
    
    # Создаем директории
    New-Item -Path "$vmPath\vhd" -ItemType Directory -Force | Out-Null
    
    # Создаем VHD
    try {
        New-VHD -Path $vhdPath -ParentPath "C:\vm\parent\parent.vhdx" -Differencing -ErrorAction Stop
        Write-Host "  VHD создан: $vhdPath" -ForegroundColor Green
    }
    catch {
        Write-Host "  Ошибка при создании VHD: $_" -ForegroundColor Red
        return $false
    }
    
    # Создаем VM
    try {
        New-VM -Name $vmName -Path $vmPath -VHDPath $vhdPath -Generation 2 -ErrorAction Stop
        Write-Host "  Виртуальная машина создана: $vmName" -ForegroundColor Green
    }
    catch {
        Write-Host "  Ошибка при создании VM: $_" -ForegroundColor Red
        return $false
    }
    
    # Настраиваем процессор и память
    try {
        Set-VMProcessor -VMName $vmName -Count $cpuCount -ErrorAction Stop
        Set-VMMemory -VMName $vmName -StartupBytes ($memoryGB * 1GB) -ErrorAction Stop
        Write-Host "  Настройки процессора ($cpuCount ядер) и памяти (${memoryGB}GB) применены" -ForegroundColor Green
    }
    catch {
        Write-Host "  Ошибка при настройке VM: $_" -ForegroundColor Red
        return $false
    }
    
    return $true
}

# Основной скрипт
Write-Host "`n=== Создание тестового стенда ===" -ForegroundColor Cyan
Write-Host "Введите 'X' на любом этапе для выхода из скрипта" -ForegroundColor Red

# 1. Запрашиваем имя стенда
Write-Host "`n=== ШАГ 1: Имя стенда ===" -ForegroundColor Yellow
Write-Host "Введите X для выхода из скрипта" -ForegroundColor Red
$name = Read-Host "Введите имя стенда (например: mail)"

if ($name -eq "X" -or $name -eq "x") {
    Write-Host "Скрипт прерван пользователем." -ForegroundColor Red
    exit
}

Write-Host "Машины будут называться: $name.dc1, $name.exchange1 и т.д." -ForegroundColor Yellow

# 2. Запрашиваем количество DC
Write-Host "`n=== ШАГ 2: Контроллеры домена ===" -ForegroundColor Yellow
$dcCount = Get-NumberInput -Prompt "Сколько нужно контроллеров домена?" -Min 1 -Max 2

# Создаём контроллеры домена
if ($dcCount -gt 0) {
    Write-Host "`n=== Создание контроллеров домена ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $dcCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.dc$i" -vmType "dc$i" -cpuCount 2 -memoryGB 8) {
            $successCount++
        }
    }
    Write-Host "Успешно создано контроллеров домена: $successCount из $dcCount" -ForegroundColor $(if ($successCount -eq $dcCount) {"Green"} else {"Yellow"})
}

# 3. Запрашиваем количество серверов Exchange
Write-Host "`n=== ШАГ 3: Серверы Exchange ===" -ForegroundColor Yellow
$exchangeCount = Get-NumberInput -Prompt "Сколько нужно серверов Exchange?" -Min 0 -Max 2

# Создаём серверы Exchange
if ($exchangeCount -gt 0) {
    Write-Host "`n=== Создание серверов Exchange ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $exchangeCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.exchange$i" -vmType "exchange$i" -cpuCount 8 -memoryGB 32) {
            $successCount++
        }
    }
    Write-Host "Успешно создано серверов Exchange: $successCount из $exchangeCount" -ForegroundColor $(if ($successCount -eq $exchangeCount) {"Green"} else {"Yellow"})
}

# 4. Запрашиваем количество серверов SQL
Write-Host "`n=== ШАГ 4: Серверы SQL ===" -ForegroundColor Yellow
$sqlCount = Get-NumberInput -Prompt "Сколько нужно серверов SQL?" -Min 0 -Max 2

# Создаём серверы SQL
if ($sqlCount -gt 0) {
    Write-Host "`n=== Создание серверов SQL ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $sqlCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.sql$i" -vmType "sql$i" -cpuCount 8 -memoryGB 32) {
            $successCount++
        }
    }
    Write-Host "Успешно создано серверов SQL: $successCount из $sqlCount" -ForegroundColor $(if ($successCount -eq $sqlCount) {"Green"} else {"Yellow"})
}

# 5. Запрашиваем количество рядовых серверов
Write-Host "`n=== ШАГ 5: Рядовые серверы ===" -ForegroundColor Yellow
$serverCount = Get-NumberInput -Prompt "Сколько нужно рядовых серверов?" -Min 0 -Max 2

# Создаём рядовые серверы
if ($serverCount -gt 0) {
    Write-Host "`n=== Создание рядовых серверов ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $serverCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.server$i" -vmType "server$i" -cpuCount 8 -memoryGB 32) {
            $successCount++
        }
    }
    Write-Host "Успешно создано рядовых серверов: $successCount из $serverCount" -ForegroundColor $(if ($successCount -eq $serverCount) {"Green"} else {"Yellow"})
}

# Итоговый отчет
Write-Host "`n=== ИТОГОВАЯ СВОДКА ===" -ForegroundColor Cyan
Write-Host "Имя стенда: $name" -ForegroundColor Yellow
Write-Host "Контроллеры домена: $dcCount" -ForegroundColor Green
Write-Host "Серверы Exchange: $exchangeCount" -ForegroundColor Green
Write-Host "Серверы SQL: $sqlCount" -ForegroundColor Green
Write-Host "Рядовые серверы: $serverCount" -ForegroundColor Green

$totalVMs = $dcCount + $exchangeCount + $sqlCount + $serverCount
Write-Host "`nВсего виртуальных машин для создания: $totalVMs" -ForegroundColor Magenta

Write-Host "`nСоздание стенда завершено!" -ForegroundColor Green
Write-Host "Скрипт завершен." -ForegroundColor Cyan