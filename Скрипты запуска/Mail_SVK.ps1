Add-Type -assembly "Microsoft.Office.Interop.Outlook"
Import-Module "W:\example\Library_Mail.ps1"
Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\Library_SVK.ps1"
Import-Module "W:\example\SendErr.ps1"

Function fProcessReports {
    
    $MainLogPath = ManageLogs

    [SVK_operations]$SVK_operations = [SVK_operations]::New()

    [bool]$bReady = $true

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

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Новый цикл обработки" | Tee-Object $MainLogPath -Append

        $bAuthOK = $SVK_operations.fNeedAuth()

        if($bAuthOK -ne $true){

            try {

                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Авторизация" | Tee-Object $MainLogPath -Append
                $bAuthOK = $SVK_operations.fGetAuth()

            }
            Catch {
                $bAuthOK = $false
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Ошибка авторизации" | Tee-Object $MainLogPath -Append
            }

        }

        if($bAuthOK) {

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Отправка и получение почты" | Tee-Object $MainLogPath -Append
            $oNamespace.SendAndReceive($false)
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Пауза 15 сек" | Tee-Object $MainLogPath -Append
            Start-Sleep -Seconds 15
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Старт" | Tee-Object $MainLogPath -Append
            Start-Sleep -Seconds 3

        }

        Write-Output "$(Get-Date) "

        if($bAuthOK) {
            
            $MainLogPath = ManageLogs
            fDoWork $oOutlook $oNamespace $oInboxFolder
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Пауза 30 секунд" | Tee-Object $MainLogPath -Append
            Start-Sleep -Seconds 30
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    !Старт" | Tee-Object $MainLogPath -Append
            Start-Sleep -Seconds 5

        }

    if($Error){

        $err = $Error.Item(0).ToString()
        if($err -notmatch "-SVK.txt"){
            
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обнаружена ошибка в процессе работы!" | Tee-Object $MainLogPath -Append
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $err" | Tee-Object $MainLogPath -Append
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Отправка письма об ошибке" | Tee-Object $MainLogPath -Append
            $SendErr.SendMail("Mail_SVK", "Ошибка обработки входящих писем MAIL_SVK", "$err")
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Письмо отправлено." | Tee-Object $MainLogPath -Append
        }
        $Error.Clear()

        }
    
    }

}

Function fDoWork {

    Param(

        $oOutlook,
        $oNamespace,
        $oInboxFolder

    )

    [Mail]$Mail = [Mail]::New()
    [string[][]]$aFormsTitle = $Mail.aFormsNotify
    [string[]]$aCBRAddresses = $Mail.aCBRAddresses
    
    [Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
    [string]$sReportsPath = $Path_to_Folders.sReportsPath
    [Run_exe]$Run_exe = [Run_exe]::New()
    [File_operations]$File_operations = [File_operations]::New()
    

    $MainLogPath = ManageLogs

    $oInboxItems = $oInboxFolder.items

    foreach($oInboxItem in $oInboxItems) {

        Start-Sleep -Seconds 3
        [string]$sInboxItemSubject = $oInboxItem.Subject.ToUpper()
        [string]$sInboxItemAddress = $oInboxItem.SenderName.ToUpper()
        [string]$sInboxItemTo = $oInboxItem.To.ToUpper()
        $itsKnowReport = $false

        for($iNum2 = 0; $iNum2 -lt $aFormsTitle.count; $iNum2++) {
            
		    if((((($sInboxItemSubject -eq ("RE:" + $aFormsTitle[$iNum2][0])) -and (($sInboxItemAddress -eq $aCBRAddresses[11]) -or ($sInboxItemAddress -eq $aCBRAddresses[0]))) -or ($sInboxItemSubject -eq ("RE: " + $aFormsTitle[$iNum2][0])))) -or ($sInboxItemSubject -eq ($aFormsTitle[$iNum2][0] + " NOTIFICATION")) -or (($sInboxItemAddress -eq $aCBRAddresses[11]) -and ($sInboxItemSubject -eq ("MPSO " + $aFormsTitle[$iNum2][0]))) -or (($sInboxItemAddress -eq $aCBRAddresses[11]) -and ($sInboxItemSubject -eq ("Re: MPSO " + $aFormsTitle[$iNum2][0])))) {
                $itsKnowReport = $true
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append
                
                if($oInboxItem.Attachments.Count -gt 0) {

                    foreach($oAttachment in $oInboxItem.Attachments) {

                        [string]$sFileName = $oAttachment.FileName
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                        [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, $aFormsTitle[$iNum2][0], 'IN')
                        
                    }
                        
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append
                    
                }#Сохранение отчетности ЦБ
                 
            } elseif(($aFormsTitle[$iNum2][0] -eq "FNS") -and ($sInboxItemAddress -eq $aCBRAddresses[7].ToUpper()) -and ($sInboxItemSubject.ToUpper().Contains("ПЕРЕДАЧА") -or $sInboxItemSubject.Contains("RE:FNS"))) {
                $itsKnowReport = $true
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append

                if($oInboxItem.Attachments.Count -gt 0) {

                    foreach($oAttachment in $oInboxItem.Attachments) {

                        [string]$sFileName = $oAttachment.FileName
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                        [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, $aFormsTitle[$iNum2][0], 'IN')
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Первичная распаковка полученных файлов" | Tee-Object $MainLogPath -Append
                        $Run_exe.fUnzipReport($sReportPath)
                        Start-Sleep -Seconds 2
                        $File_operations.fDelFile($sReportPath)
                        
                    }
                        
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append

                }#Сохранение и первичная распаковка 440

            } elseif(($aFormsTitle[$iNum2][0] -eq "4512") -and ($sInboxItemAddress -eq $aCBRAddresses[15].ToUpper()) -and ($sInboxItemSubject.ToUpper().Contains("ПЕРЕДАЧА") -or $sInboxItemSubject.Contains("RE:4512"))) {
                $itsKnowReport = $true
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append

                if($oInboxItem.Attachments.Count -gt 0) {

                    foreach($oAttachment in $oInboxItem.Attachments) {

                        [string]$sFileName = $oAttachment.FileName
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                        [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, $aFormsTitle[$iNum2][0], 'IN')
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Первичная распаковка полученных файлов" | Tee-Object $MainLogPath -Append
                        $Run_exe.fUnzipReport($sReportPath)
                        Start-Sleep -Seconds 2
                        $File_operations.fDelFile($sReportPath)
                        
                    }
                        
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append

                }#Сохранение и первичная распаковка 4512

            } elseif(($aFormsTitle[$iNum2][0] -eq "FSFM550") -and ($sInboxItemAddress -eq $aCBRAddresses[10].ToUpper()) -and ($sInboxItemSubject.ToUpper().Contains("ПЕРЕДАЧА АРХИВНЫХ ФАЙЛОВ") -or $sInboxItemSubject.Contains("RE:FSFM550"))) {
                $itsKnowReport = $true
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append

                if($oInboxItem.Attachments.Count -gt 0) {

                    foreach($oAttachment in $oInboxItem.Attachments) {

                        [string]$sFileName = $oAttachment.FileName
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                        [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, $aFormsTitle[$iNum2][0], 'IN')
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Первичная распаковка полученных файлов" | Tee-Object $MainLogPath -Append
                        $Run_exe.fUnzipReport($sReportPath)
                        Start-Sleep -Seconds 2
                        $File_operations.fDelFile($sReportPath)
                                                
                    }
                        
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append

                }#Сохранение и первичная распаковка 550

            } elseif(($aFormsTitle[$iNum2][0] -eq "MIFNS") -and ((($sInboxItemSubject -eq "311-П") -or ($sInboxItemSubject -eq "311-П.") -or $sInboxItemSubject.ToUpper().Contains("RE:MIFNS")))) {
                $itsKnowReport = $true
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append

                if($oInboxItem.Attachments.Count -gt 0) {

                    foreach($oAttachment in $oInboxItem.Attachments) {

                        [string]$sFileName = $oAttachment.FileName
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                        [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, $aFormsTitle[$iNum2][0], 'IN')
                        
                    }
                        
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма $($aFormsTitle[$iNum2][0])" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append

                }#Сохранение MIFNS

            } elseif(($aFormsTitle[$iNum2][0] -eq "NBKI") -and ($sInboxItemAddress -eq $aCBRAddresses[12])) {
                $itsKnowReport = $true
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы NBKI" | Tee-Object $MainLogPath -Append

                if($oInboxItem.Attachments.Count -gt 0) {

                    foreach($oAttachment in $oInboxItem.Attachments) {

                        [string]$sFileName = $oAttachment.FileName
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                        [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, "NBKI", 'IN')
                        
                    }
                        
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма NBKI" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append

                }#Сохранение NBKI

            } elseif(($aFormsTitle[$iNum2][0] -match "VBK") -and ($sInboxItemAddress -eq $aCBRAddresses[13]) -and $sInboxItemSubject.Contains("VBK")) {
                $itsKnowReport = $true
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка сохранения вложения формы VBK" | Tee-Object $MainLogPath -Append

                if($oInboxItem.Attachments.Count -gt 0) {

                    foreach($oAttachment in $oInboxItem.Attachments) {

                        [string]$sFileName = $oAttachment.FileName
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сохранение из письма вложения $sFileName" | Tee-Object $MainLogPath -Append
                        [string]$sReportPath = $Mail.fSaveAttachment($sReportsPath, $oAttachment, "VBK", 'IN')
                        
                    }
                                        
                $Mail.fMoveReportMessage($oInboxFolder, $oInboxItem, $aFormsTitle[$iNum2][0])
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступившего письма VBK" | Tee-Object $MainLogPath -Append
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    " | Tee-Object $MainLogPath -Append

                }#Сохранение VBK

            }

        }

        if((($sInboxItemTo -eq $Mail.sBankSVKAddress.ToUpper()) -or ($sInboxItemTo -eq $Mail.sBankSVKAddressReserv.ToUpper())) -and ($itsKnowReport -eq $false)){
                
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Поступило неизвестное письмо с заголовком: $sInboxItemSubject с адреса $sInboxItemAddress" | Tee-Object $MainLogPath -Append
            $to = $SendErr.to
            $subject = "SVK - Поступило неизвестное письмо"
            $body = "Поступило неизвестное письмо с заголовком: $sInboxItemSubject с адреса: $sInboxItemAddress. Письмо перемещено в папку Входящие - N2 почтового ящика example"
            $PathToSaveMSG = "$logDir\$($oInboxItem.SenderName.Split("@")[0])-$(Get-Date -Format "yyyyMMddHHmmss").msg"
            $pathToMsg = $Mail.fSaveMSG($oInboxItem, $PathToSaveMSG)
            $SendErr.SendMSG($pathToMsg, $subject, $body)
            $File_operations.fDelFile($pathToMsg)
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Отправка информации о неизвестном письме на адрес $to" | Tee-Object $MainLogPath -Append
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос входящего в папку N2" | Tee-Object $MainLogPath -Append
            $Mail.fMoveUnknowMessage($oInboxFolder, $oInboxItem)

        } #Сообщение о неизвестной отчетности 


    } #Сохранение входящих из ЦБ в папки

    $sReportsPaths = $Path_to_Folders.sReportsPath + "\"
    $sReportsPaths = Get-ChildItem $sReportsPaths -Directory

    foreach($sFoldersForms in $sReportsPaths){

        $sFoldersFormsOUT = $sFoldersForms.FullName + "\OUT\FORSEND"
        
        [string]$FolderName = $sFoldersForms.Name
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Поиск файлов на отправку по форме $FolderName"
        $sFilesInFolder = Get-ChildItem $sFoldersFormsOUT -File

        if($sFilesInFolder -ne $null){

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки исходящих файлов формы $FolderName" | Tee-Object $MainLogPath -Append

            foreach($sFileInFolder in $sFilesInFolder){
                
                [string]$sForm = $sFoldersForms.Name.ToUpper()
                if($sForm -eq "FNS"){
                    
                    $sMailAddress = $aCBRAddresses[7] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                
                } elseif($sForm -eq "FSFM550"){
                
                    $sMailAddress = $aCBRAddresses[10] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                
                } elseif($sForm -eq "KFM"){
                
                    $sMailAddress = $aCBRAddresses[1] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                
                } elseif($sForm -eq "4512"){
                
                    $sMailAddress = $aCBRAddresses[15] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                
                } elseif($sForm -eq "VBK"){
                
                    $sMailAddress = $aCBRAddresses[13] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                
                } elseif($sForm -eq "MIFNS"){
                
                    $sMailAddress = $aCBRAddresses[2] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                
                } elseif(($sForm -ne "NBKI") -and ($sFileInFolder.Extension.ToUpper() -eq ".XML")){
                    
                    $sMailAddress = $aCBRAddresses[11] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                    SendConfirm $sForm $aFormsTitle

                }elseif(($sForm -ne "NBKI") -and ($sFileInFolder.Extension.ToUpper() -ne ".XML")){
                    
                    $sMailAddress = $aCBRAddresses[0] # example@example.ru
                    Send $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                    SendConfirm $sForm $aFormsTitle

                }elseif($sForm -eq "NBKI"){
                    $sForm = "1501BB000001"
                    $sMailAddress = $aCBRAddresses[14] # example@example.ru
                    Send_NBKI $sFileInFolder $oOutlook $sMailAddress $sForm $File_operations
                
                }
            
            }
        
        }


    } #Отправка в ЦБ


}

function Send {

    Param(
    
        $sFileInFolder,
        $oOutlook,
        $sMailAddress,
        $sForm,
        $File_operations
    
    )

    $sReportPath = $sFileInFolder.FullName
    $sFileName = $sFileInFolder.Name
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Отправка в ЦБ файла $sFileName" | Tee-Object $MainLogPath -Append
    $Mail.fSendReportFile($oOutlook, $sReportPath, $sForm, $sMailAddress)
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $sFileName" | Tee-Object $MainLogPath -Append
    $File_operations.fDelFile($sReportPath)
}

function Send_NBKI {

    Param(
    
        $sFileInFolder,
        $oOutlook,
        $sMailAddress,
        $sForm,
        $File_operations
    
    )

    $sReportPath = $sFileInFolder.FullName
    $sFileName = $sFileInFolder.Name
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Отправка в NBKI файла $sFileName" | Tee-Object $MainLogPath -Append
    $Mail.fSendNBKIFile($oOutlook, $sReportPath, $sForm, $sMailAddress)
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $sFileName" | Tee-Object $MainLogPath -Append
    $File_operations.fDelFile($sReportPath)
}

function ManageLogs{
    
    $logFileName = $logDir + "\$(Get-Date -Format 'yyyy-MM-dd')-SVK.txt"

    if(Test-Path $logFileName){
    
        return $logFileName
    
    } else {

        $logfile = New-Item -Path $logFileName -ItemType File -ErrorAction Ignore

        return $logfile.FullName

    }

}

function SendConfirm{
    Param(
        $sFormName,
        $aFormsTitle
    )


    $MailGroupsName = $($aFormsTitle | Where-Object {$_ -eq "$sFormName"})[1]
    $MailGroups = $MailGroupsName.Split("#")
    foreach($MailGroup in $MailGroups){

        $to = "ROReports" + $MailGroup + "@resocreditbank.ru"
        $subject = $sFormName + " успешно отправлено!"
        $body = "$sFormName отправлено в $(Get-Date)"
        $SendErr.SendConfirm($to, $subject, $body)
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Отправка подтверждения на адрес $to" | Tee-Object $MainLogPath -Append

    }

}

[SendErrors]$SendErr = [SendErrors]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
$logDir = $Path_to_Folders.sMailLogs + "\Mail_SVK"
$MainLogPath = ManageLogs

fProcessReports