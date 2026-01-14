<#
.SYNOPSIS
    Инструмент очистки почтовых ящиков Exchange
.DESCRIPTION
    Этот скрипт выполняет поиск и удаление писем из всех почтовых ящиков Exchange по критерию темы или отправителя.
    Используется двухэтапный процесс: сначала показывается предварительный просмотр найденных писем, затем выполняется удаление после подтверждения.
    
    ОСОБЕННОСТИ:
    - Поиск во всех почтовых ящиках по заданным критериям
    - Предварительный просмотр с оценкой количества найденных писем
    - Требуется подтверждение пользователя перед удалением
    - Оптимизированное удаление только из ящиков с найденными письмами
    - Индикатор прогресса при удалении
    
.PARAMETER None
    Все параметры собираются интерактивно во время выполнения скрипта.
    
.EXAMPLE
    PS> .\Remove-EmailsFromAllMailboxes.ps1
    
    Скрипт запросит:
    1. Учетные данные администратора Exchange
    2. Имя сервера Exchange
    3. Критерий поиска (Тема или Отправитель)
    4. Термин для поиска
    5. Подтверждение перед удалением
    
.NOTES
    Автор: Александр Серкин
    Версия: 2.0
    Дата: 14.01.2026
    
    НЕОБХОДИМЫЕ ПРАВА:
    - Роль администратора Exchange или права на поиск в почтовых ящиках
    - Возможность выполнения командлета Search-Mailbox
    
    ВНИМАНИЕ:
    - Этот скрипт НАВСЕГДА удаляет письма
    - Всегда тестируйте в непродуктивной среде
    - Убедитесь, что у вас есть резервные копии
    - Проверяйте критерии поиска перед подтверждением удаления
    
.LINK
    https://docs.microsoft.com/ru-ru/powershell/module/exchange/search-mailbox
#>

#region Подключение к Exchange
Write-Host "=== Инструмент очистки почтовых ящиков Exchange ===" -ForegroundColor Cyan
Write-Host "Введите учетные данные с правами поиска в почтовых ящиках Exchange" -ForegroundColor Yellow
$UserCredential = Get-Credential

$mailserver = Read-Host "Введите имя вашего сервера Exchange"
Write-Host "Подключение к серверу Exchange: $mailserver..." -ForegroundColor Cyan

try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$mailserver/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction Stop
    Import-PSSession $Session -DisableNameChecking -ErrorAction Stop
    Write-Host "Успешное подключение к Exchange" -ForegroundColor Green
}
catch {
    Write-Host "Ошибка подключения к Exchange: $_" -ForegroundColor Red
    exit 1
}
#endregion

#region Главное меню
Write-Host "`n=== ВЫБЕРИТЕ КРИТЕРИЙ ПОИСКА ===" -ForegroundColor Yellow
Write-Host "1. Поиск по теме письма"
Write-Host "2. Поиск по адресу отправителя"

do {
    $choice = Read-Host "Введите номер опции (1-2)"
    
    switch ($choice) {
        "1" {
            # Поиск по теме
            $subject = Read-Host "Введите тему письма (частично или полностью)"
            if (-not [string]::IsNullOrWhiteSpace($subject)) {
                Write-Host "Поиск писем с темой, содержащей: '$subject'..." -ForegroundColor Cyan
                
                # Фаза 1: Предварительный просмотр результатов
                Write-Host "`n=== ПРЕДВАРИТЕЛЬНЫЙ ПРОСМОТР ===" -ForegroundColor Yellow
                Write-Host "Сканирование всех почтовых ящиков..." -ForegroundColor Cyan
                
                try {
                    $previewResults = Get-Mailbox -ResultSize Unlimited | Search-Mailbox -SearchQuery "subject:`"$subject*`"" -EstimateResultOnly -WarningAction SilentlyContinue -ErrorAction Stop
                    
                    # Фильтрация результатов - только ящики с найденными письмами
                    $mailboxesWithItems = $previewResults | Where-Object { $_.ResultItemsCount -ge 1 }
                    
                    Write-Host "Поиск завершен." -ForegroundColor Green
                    
                    if ($mailboxesWithItems.Count -gt 0) {
                        Write-Host "`nНайдены письма в следующих почтовых ящиках:" -ForegroundColor Green
                        $mailboxesWithItems | Select-Object @{
                            Name = "ИмяЯщика"
                            Expression = { ($_.Identity.ToString().Split('/')[-1]) }
                        }, ResultItemsCount | Format-Table -AutoSize
                        
                        $totalFound = ($mailboxesWithItems | Measure-Object -Property ResultItemsCount -Sum).Sum
                        Write-Host "Всего найдено писем: $totalFound в $($mailboxesWithItems.Count) ящике(ах)" -ForegroundColor Green
                        
                        # Запрос подтверждения на удаление
                        Write-Host "`n=== ПОДТВЕРЖДЕНИЕ УДАЛЕНИЯ ===" -ForegroundColor Red
                        Write-Host "ВНИМАНИЕ: Письма будут удалены НАВСЕГДА!" -ForegroundColor Red
                        $confirm = Read-Host "Вы действительно хотите удалить эти письма? (Да/Нет)"
                        
                        if ($confirm -eq "Да" -or $confirm -eq "да" -or $confirm -eq "y" -or $confirm -eq "Y") {
                            Write-Host "`nЗапуск процесса удаления..." -ForegroundColor Cyan
                            
                            # Фаза 2: Выполнение удаления
                            $deleteResults = @()
                            $deletedCount = 0
                            $totalToDelete = $mailboxesWithItems.Count
                            
                            # Показываем прогресс при удалении
                            foreach ($mailbox in $mailboxesWithItems) {
                                $deletedCount++
                                Write-Progress -Activity "Удаление писем" -Status "Обработка почтовых ящиков" `
                                    -PercentComplete (($deletedCount / $totalToDelete) * 100) `
                                    -CurrentOperation "$deletedCount из $totalToDelete"
                                
                                try {
                                    $result = Search-Mailbox -Identity $mailbox.Identity -SearchQuery "subject:`"$subject*`"" -DeleteContent -Force -WarningAction SilentlyContinue -ErrorAction Stop
                                    
                                    if ($result) {
                                        $deleteResults += @{
                                            ИмяЯщика = ($mailbox.Identity.ToString().Split('/')[-1])
                                            КоличествоУдаленных = $result.ResultItemsCount
                                        }
                                    }
                                }
                                catch {
                                    Write-Host "Ошибка удаления из $($mailbox.Identity): $_" -ForegroundColor Yellow
                                }
                            }
                            
                            Write-Progress -Activity "Удаление писем" -Completed
                            
                            if ($deleteResults.Count -gt 0) {
                                Write-Host "`n=== РЕЗУЛЬТАТЫ УДАЛЕНИЯ ===" -ForegroundColor Green
                                Write-Host "Успешно удалены письма из почтовых ящиков:" -ForegroundColor Green
                                $deleteResults | Select-Object ИмяЯщика, КоличествоУдаленных | Format-Table -AutoSize
                                
                                $totalDeleted = ($deleteResults | Measure-Object -Property КоличествоУдаленных -Sum).Sum
                                Write-Host "Всего удалено писем: $totalDeleted из $($deleteResults.Count) ящика(ов)" -ForegroundColor Green
                            } else {
                                Write-Host "Письма не были удалены." -ForegroundColor Yellow
                            }
                        } else {
                            Write-Host "Удаление отменено пользователем." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "Писем с указанной темой не найдено." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Ошибка при поиске: $_" -ForegroundColor Red
                }
                
                $validChoice = $true
            } else {
                Write-Host "Тема не может быть пустой!" -ForegroundColor Red
                $validChoice = $false
            }
        }
        
        "2" {
            # Поиск по отправителю
            $from = Read-Host "Введите адрес отправителя (частично или полностью)"
            if (-not [string]::IsNullOrWhiteSpace($from)) {
                Write-Host "Поиск писем от отправителя, содержащего: '$from'..." -ForegroundColor Cyan
                
                # Фаза 1: Предварительный просмотр результатов
                Write-Host "`n=== ПРЕДВАРИТЕЛЬНЫЙ ПРОСМОТР ===" -ForegroundColor Yellow
                Write-Host "Сканирование всех почтовых ящиков..." -ForegroundColor Cyan
                
                try {
                    $previewResults = Get-Mailbox -ResultSize Unlimited | Search-Mailbox -SearchQuery "from:`"$from`"" -EstimateResultOnly -WarningAction SilentlyContinue -ErrorAction Stop
                    
                    # Фильтрация результатов - только ящики с найденными письмами
                    $mailboxesWithItems = $previewResults | Where-Object { $_.ResultItemsCount -ge 1 }
                    
                    Write-Host "Поиск завершен." -ForegroundColor Green
                    
                    if ($mailboxesWithItems.Count -gt 0) {
                        Write-Host "`nНайдены письма в следующих почтовых ящиках:" -ForegroundColor Green
                        $mailboxesWithItems | Select-Object @{
                            Name = "ИмяЯщика"
                            Expression = { ($_.Identity.ToString().Split('/')[-1]) }
                        }, ResultItemsCount | Format-Table -AutoSize
                        
                        $totalFound = ($mailboxesWithItems | Measure-Object -Property ResultItemsCount -Sum).Sum
                        Write-Host "Всего найдено писем: $totalFound в $($mailboxesWithItems.Count) ящике(ах)" -ForegroundColor Green
                        
                        # Запрос подтверждения на удаление
                        Write-Host "`n=== ПОДТВЕРЖДЕНИЕ УДАЛЕНИЯ ===" -ForegroundColor Red
                        Write-Host "ВНИМАНИЕ: Письма будут удалены НАВСЕГДА!" -ForegroundColor Red
                        $confirm = Read-Host "Вы действительно хотите удалить эти письма? (Да/Нет)"
                        
                        if ($confirm -eq "Да" -or $confirm -eq "да" -or $confirm -eq "y" -or $confirm -eq "Y") {
                            Write-Host "`nЗапуск процесса удаления..." -ForegroundColor Cyan
                            
                            # Фаза 2: Выполнение удаления
                            $deleteResults = @()
                            $deletedCount = 0
                            $totalToDelete = $mailboxesWithItems.Count
                            
                            # Показываем прогресс при удалении
                            foreach ($mailbox in $mailboxesWithItems) {
                                $deletedCount++
                                Write-Progress -Activity "Удаление писем" -Status "Обработка почтовых ящиков" `
                                    -PercentComplete (($deletedCount / $totalToDelete) * 100) `
                                    -CurrentOperation "$deletedCount из $totalToDelete"
                                
                                try {
                                    $result = Search-Mailbox -Identity $mailbox.Identity -SearchQuery "from:`"$from`"" -DeleteContent -Force -WarningAction SilentlyContinue -ErrorAction Stop
                                    
                                    if ($result) {
                                        $deleteResults += @{
                                            ИмяЯщика = ($mailbox.Identity.ToString().Split('/')[-1])
                                            КоличествоУдаленных = $result.ResultItemsCount
                                        }
                                    }
                                }
                                catch {
                                    Write-Host "Ошибка удаления из $($mailbox.Identity): $_" -ForegroundColor Yellow
                                }
                            }
                            
                            Write-Progress -Activity "Удаление писем" -Completed
                            
                            if ($deleteResults.Count -gt 0) {
                                Write-Host "`n=== РЕЗУЛЬТАТЫ УДАЛЕНИЯ ===" -ForegroundColor Green
                                Write-Host "Успешно удалены письма из почтовых ящиков:" -ForegroundColor Green
                                $deleteResults | Select-Object ИмяЯщика, КоличествоУдаленных | Format-Table -AutoSize
                                
                                $totalDeleted = ($deleteResults | Measure-Object -Property КоличествоУдаленных -Sum).Sum
                                Write-Host "Всего удалено писем: $totalDeleted из $($deleteResults.Count) ящика(ов)" -ForegroundColor Green
                            } else {
                                Write-Host "Письма не были удалены." -ForegroundColor Yellow
                            }
                        } else {
                            Write-Host "Удаление отменено пользователем." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "Писем от указанного отправителя не найдено." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Ошибка при поиске: $_" -ForegroundColor Red
                }
                
                $validChoice = $true
            } else {
                Write-Host "Адрес отправителя не может быть пустым!" -ForegroundColor Red
                $validChoice = $false
            }
        }
        
        default {
            Write-Host "Неверная опция! Пожалуйста, введите 1 или 2." -ForegroundColor Red
            $validChoice = $false
        }
    }
} while ($validChoice -ne $true)
#endregion

#region Завершение работы
# Закрытие сессии Exchange
if ($Session) {
    try {
        Remove-PSSession $Session
        Write-Host "`nСессия Exchange успешно закрыта." -ForegroundColor Green
    }
    catch {
        Write-Host "Предупреждение: Не удалось корректно закрыть сессию Exchange." -ForegroundColor Yellow
    }
}

Write-Host "`n=== Выполнение скрипта завершено ===" -ForegroundColor Cyan
#endregion