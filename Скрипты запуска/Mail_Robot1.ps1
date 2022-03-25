Add-Type -assembly "Microsoft.Office.Interop.Outlook"
Import-Module "W:\example\Library_Mail.ps1"
Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

Function CheckFileExist{

    param(
    
        $sReportPath,
        $report
    
    )

    $path_to_report = $sReportPath + "\" + $report + "\" + $report
    $exist = Test-Path $path_to_report

    Return $exist

}

Function fProcessReports {

    $MainLogPath = ManageLogs

    [bool]$bReady = $true

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    ЗАПУСК ПРОЦЕССА ОБРАБОТКИ" | Tee-Object $MainLogPath -Append
    
    try {

        $oOutlook = New-Object -com Outlook.Application
        $oOutlookProcess = Get-Process Outlook
        $oNamespace = $oOutlook.GetNameSpace("MAPI")
        $oInboxFolder = $oNamespace.GetDefaultFolder(6)
        $oExplorer = $oInboxFolder.GetExplorer()

    } #Попытка создания объекта Outlook и получения папок
    Catch {

        $bReady = $false
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Ошибка создания объекта Outlook" | Tee-Object $MainLogPath -Append

    } #Ошибка создания объекта Outlook - информация об ошибке $bReady = $false

    :loop1 while($bReady) {
        $MainLogPath = ManageLogs
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Новый цикл обработки" | Tee-Object $MainLogPath -Append
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Отправка и получение почты" | Tee-Object $MainLogPath -Append
        $oNamespace.SendAndReceive($false)
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Пауза 15 сек" | Tee-Object $MainLogPath -Append
        Start-Sleep -Seconds 15
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Старт" | Tee-Object $MainLogPath -Append
        Start-Sleep -Seconds 3
        
        Write-Output "$(Get-Date) "
        fDoWork $oOutlook $oNamespace $oInboxFolder
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Пауза 30 секунд" | Tee-Object $MainLogPath -Append
        Start-Sleep -Seconds 30
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Старт" | Tee-Object $MainLogPath -Append
        Start-Sleep -Seconds 5
        
        if($Error){

        $err = $Error.Item(0).ToString()
        if($err -notmatch "-ROBOT1.txt"){
            
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обнаружена ошибка в процессе работы!" | Tee-Object $MainLogPath -Append
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $err" | Tee-Object $MainLogPath -Append
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Отправка письма об ошибке" | Tee-Object $MainLogPath -Append
            $SendErr.SendMail("Mail_SVK", "Ошибка обработки входящих писем MAIL_ROBOT1", "$err")
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Письмо отправлено." | Tee-Object $MainLogPath -Append

        }
        $Error.Clear()

        }

    } #Бесконечный цикл отправки и получения

} #Создание объекта Outlook и запуск цикла опроса потового ящика

Function fDoWork {

    Param(

        $oOutlook,
        $oNamespace,
        $oInboxFolder

    )

    [Mail]$Mail = [Mail]::New()
    [string[][]]$aFormsTitle = $Mail.aFormsNotify
    
    [File_operations]$File_operations = [File_operations]::New()
    
    [Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
    [string]$sReportsPath = $Path_to_Folders.sReportsPath

    [Run_exe]$Run_exe = [Run_exe]::New()

    $MainLogPath = ManageLogs

    $sPS = $Run_exe.sPS

    $oInboxItems = $oInboxFolder.items

    foreach($oInboxItem in $oInboxItems) {

        Start-Sleep -Seconds 3
        [string]$sInboxItemSubject = $oInboxItem.Subject.ToUpper()
        [string]$sInboxItemAddress = $oInboxItem.SenderName.ToUpper()

        if(($oInboxItem.To.Contains("ROBOT1") -and ($sInboxItemSubject -match "550")) -or ($oInboxItem.CC.Contains("example")) -and ($sInboxItemSubject -match "550")){
            
            $exsit = CheckFileExist $sReportsPath "FSFM550"
            if($exsit -eq $false){
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки на отправку 550-П" | Tee-Object $MainLogPath -Append
                [string]$RunWork550 = $sReportsPath + "\example\FSFM550_OUT.ps1"
                Start-Process -FilePath $sPS $RunWork550
                $iNum2 = 5
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма 550-П" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append
            } else {
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сообщение $sInboxItemSubject поставлено в очередь!" | Tee-Object $MainLogPath -Append
            }
                    
        } elseif(($oInboxItem.To.Contains("example") -and ($sInboxItemSubject -match "MIFNS")) -or ($oInboxItem.CC.Contains("example")) -and ($sInboxItemSubject -match "MIFNS")){

            $exsit = CheckFileExist $sReportsPath "MIFNS"
            if($exsit -eq $false){
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки на отправку MIFNS" | Tee-Object $MainLogPath -Append
                [string]$RunWorkMIFNS = $sReportsPath + "\example\MIFNS_OUT.ps1"
                Start-Process -FilePath $sPS $RunWorkMIFNS
                $iNum2 = 74
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма MIFNS" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append
            } else {
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сообщение $sInboxItemSubject поставлено в очередь!" | Tee-Object $MainLogPath -Append
            }
                    
        } elseif(($oInboxItem.To.Contains("ROBOT1") -and ($sInboxItemSubject -match "NBKI")) -or ($oInboxItem.CC.Contains("example")) -and ($sInboxItemSubject -match "NBKI")){

            $exsit = CheckFileExist $sReportsPath "NBKI"
            if($exsit -eq $false){
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки на отправку NBKI" | Tee-Object $MainLogPath -Append
                [string]$RunWorkNBKI = $sReportsPath + "\example\NBKI_OUT.ps1"
                Start-Process -FilePath $sPS $RunWorkNBKI
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[105][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма NBKI" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append
            } else {
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сообщение $sInboxItemSubject поставлено в очередь!" | Tee-Object $MainLogPath -Append
            }
                    
        } elseif(($oInboxItem.To.Contains("example")) -or ($oInboxItem.CC.Contains("example"))) {

            for($iNum2 = 0; $iNum2 -lt $aFormsTitle.count; $iNum2++) {
            
		        if(($sInboxItemSubject -eq $aFormsTitle[$iNum2][0]) -or ($sInboxItemSubject -eq ($aFormsTitle[$iNum2][0] + "XML").ToUpper())) {
                    
                    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append

                    if($oInboxItem.Attachments.Count -gt 0) {

                        foreach($oAttachment in $oInboxItem.Attachments) {

                            [string]$sFileName = $oAttachment.FileName
                            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                            [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, $aFormsTitle[$iNum2][0], 'OUT')
                        
                        }
                        
                    $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append
                    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append
                    
                    }
                 
                } 

            }

        }

    } #Сохранение вложения в папку /№отчетности/OUT

    $sReportsPaths = $Path_to_Folders.sReportsPath
    $sReportsPaths = Get-ChildItem $sReportsPaths -Directory

    foreach($sFoldersForms in $sReportsPaths){
        
        [string[][]]$sFormNotify = $Mail.aFormsNotify

        $sFoldersFormsIn = $sFoldersForms.FullName + "\IN\INBOX"
        $sFilesInFolder = Get-ChildItem $sFoldersFormsIn -File

        if($sFilesInFolder -ne $null){
        [string]$FolderName = $sFoldersForms.Name
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки входящих файлов формы $FolderName" | Tee-Object $MainLogPath -Append

            foreach($sFileInFolder in $sFilesInFolder){
                
                [string]$sForm = $sFoldersForms.Name.ToUpper()
                [string[][]]$sMailReportsRole = $sFormNotify -match $sForm
                $sReportsRole = $sMailReportsRole[0][1]
                $sReplyFile = $sFileInFolder.FullName
                $sFileName = $sFileInFolder.Name
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Отправка оператору полученного файла $sFileName" | Tee-Object $MainLogPath -Append
                $Mail.fSendReplyFile($oOutlook, $sReplyFile, $sForm, $sReportsRole)
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки файла $sFileName" | Tee-Object $MainLogPath -Append
                $File_operations.fDelFile($sReplyFile)
            
            }
        
        }


    } #Отправка квитанций на почту


} 

function ManageLogs{
    
    $logFileName = $logDir + "\$(Get-Date -Format 'yyyy-MM-dd')-ROBOT1.txt"

    if(Test-Path $logFileName){
    
        return $logFileName
    
    } else {

        $logfile = New-Item -Path $logFileName -ItemType File -ErrorAction Ignore

        return $logfile.FullName

    }

}

[SendErrors]$SendErr = [SendErrors]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
$logDir = $Path_to_Folders.sMailLogs + "\Mail_Robot1"
$MainLogPath = ManageLogs

fProcessReports