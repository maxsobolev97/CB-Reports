Import-Module "W:\example\SendErr.ps1"

[SendErrors]$SendErrors = [SendErrors]::new()

While($true){
    
    Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Поиск ошибок CrypTool"
    $proc = Get-Process "CrypTool" -ErrorAction SilentlyContinue
    $status = $proc.MainWindowTitle

    if($status -eq "Ошибка"){

        $SendErrors.SendMail("CrypTool", "Ошибка работы CrypTool", "Обнаружена ошибка в работе CrypTool. Окно закрыто, обработка продолжена, ошибочные файлы перемещены в папку ERR.")
        $proc.CloseMainWindow()

    }

    [System.GC]::Collect()
    Start-Sleep 10

}