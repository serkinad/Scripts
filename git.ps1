<#
.SYNOPSIS
    Displays text in green color / Выводит текст зелёным цветом

.DESCRIPTION
    This script takes a text string and displays it in the console in green color.
    Useful for highlighting important messages, successful operations or notifications.
    
    Этот скрипт принимает текстовую строку и отображает её в консоли зелёным цветом.
    Полезно для выделения важных сообщений, успешных операций или уведомлений.

.PARAMETER Text
    Text string to display in console / Текстовая строка для вывода в консоль

.PARAMETER NoNewLine
    Do not add a new line after text (similar to Write-Host -NoNewLine)
    Не добавлять перевод строки после текста (аналогично Write-Host -NoNewLine)

.EXAMPLE
    .\ColorText.ps1 -Text "Operation completed successfully!"
    Displays: Operation completed successfully! (in green color)
    Выводит: Operation completed successfully! (зелёным цветом)

.EXAMPLE
    .\ColorText.ps1 "Внимание! Проверьте настройки"
    Displays: Внимание! Проверьте настройки (in green color)
    Выводит: Внимание! Проверьте настройки (зелёным цветом)

.NOTES
    Author: Ваше Имя / Your Name
    Version: 1.1
    Created: 2024-01-15

.LINK
    https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-host
#>


param(
    [Parameter(Mandatory=$true)]
    [string]$text
)

Write-Host $text -ForegroundColor Green