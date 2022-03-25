Class Mail{
    
    [string]$aAdminAdrress = "example@example.ru" 
    [string]$sBankSVKAddress = "example@example.ru"
    [string]$sBankSVKAddressReserv = "example@example.ru"
    [string]$sBankNBKIAddress = "example@example.ru"

    [string[]]$aCBRAddresses = @("example@ext-gate.svk.mskgtu.cbr.ru",
                       "example@example",
                       "example@example",
                       "example@example",
                       "example@example",
                       "example@example",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru",
                       "example@example.ru") #Адреса ЦБ

    $aFormsNotify = @(
                        @("BALANCE","Acc#FPA"),
                        @("BALANCEDAY","Acc#FPA"),
                        @("KFM","FM"),
                        @("ELECS","Elecs"),
                        @("FNS","FNS"),
                        @("FSFM550","FSFM550"),
                        @("F024","Acc"),
                        @("F051","Law"),
                        @("F110","Rep"),
                        @("F115","Rep"),
                        @("K115","FPA#Rep"),
                        @("F116","Rep"),
                        @("F117","Rep"),
                        @("F118","Rep"),
                        @("K118","FPA#Rep"),
                        @("F119","Rep"),
                        @("F120","Rep"),
                        @("F125","Rep"),
                        @("F126","Rep"),
                        @("F127","Rep"),
                        @("F128","Rep"),
                        @("F129","Rep"),
                        @("F136","Acc"),
                        @("F155","Rep"),
                        @("F157","Rep"),
                        @("F159","Oper"),
                        @("F162T","SVK"),
                        @("F171","Law"),
                        @("F193T","FM"),
                        @("F202","Oper"),
                        @("F203","Rep"),
                        @("F250","Rep"),
                        @("F251","Oper"),
                        @("F255","Back"),
                        @("F257","Rep"),
                        @("F258","Oper"),
                        @("F259","Oper"),
                        @("F260","Oper"),
                        @("F302","Rep"),
                        @("F303","Rep"),
                        @("F316","Rep"),
                        @("F345","Rep"),
                        @("F345P","Law"),
                        @("F350","Oper"),
                        @("F401","Cur#FPA")
                        @("F402","Cur"),
                        @("F404","Cur"),
                        @("F405","Cur"),
                        @("F406","Cur"),
                        @("F407","Cur"),
                        @("F410","Cur"),
                        @("F501F603","Rep"),
                        @("F601","Cur"),
                        @("F608","Cur"),
                        @("F634","FPA"),
                        @("F639","Rep"),
                        @("F652","Cur"),
                        @("D664","Cur"),
                        @("F664","Cur"),
                        @("F665","Cur"),
                        @("F702","Rep"),
                        @("F707","Rep"),
                        @("F711","Rep"),
                        @("F801","Law"),
                        @("F801-F805","FPA"),
                        @("K808","FPA"),
                        @("K813","FPA"),
                        @("F815","Treasury"),
                        @("F816","Treasury"),
                        @("F817","Treasury"),
                        @("F818","Treasury"),
                        @("F906-F908","Acc"),
                        @("F909","Rep"),
                        @("F7504","Law"),
                        @("FISV","Acc"),
                        @("MIFNS","Oper"),
                        @("PUBLIC","FPA"),
                        @("SPFS","Rep#SVK"),
                        @("OPDS","Acc"),
                        @("CBR1","SVK"),
                        @("CBR2","SVK"),
                        @("CBR3","SVK"),
                        @("F364P","Cur"),
                        @("F402P","Cur"),
                        @("MIR","Rep"),
                        @("TPPU","SVK"),
                        @("CIK","FM"),
                        @("F706","FPA#Rep"),
                        @("F708","FPA#Rep"),
                        @("K127","FPA#Rep"),
                        @("INRKO","FPA"),
                        @("MSPRIM","Acc"),
                        @("F070","FPA"),
                        @("F0409203","Oper"),
                        @("PZAP","FPA"),
                        @("INRBG","FPA"),
                        @("F111","FPA"),
                        @("MSAZ","Acc"),
                        @("EFUDKO","Law"),
                        @("F604","FPA"),
                        @("F610","Acc#FPA"),
                        @("AUDIT","FPA"),
                        @("MSG1","Acc#Rep"),
                        @("F704","Rep"),
                        @("F704","Rep"),
                        @("F712","Rep"),
                        @("NBKI","Rep"),
                        @("F0403230","Rep"),
                        @("F0403231","Rep"),
                        @("F0403232","Rep"),
                        @("F0403233","Rep"),
                        @("VBK","Cur"),
                        @("F910","Rep"),
                        @("F345D","Rep"),
                        @("4512","Cur"),
                        @("V664","Cur#Rep")
                     ) #Формы отчетности и группы рассылки

    fMoveReportMessage($oInboxFolder, $oInboxItem, $sForm) {

        $oTargetFolder = $oInboxFolder.Folders.Item("Reports").Folders.Item($sForm)
        [void]$oInboxItem.Move($oTargetFolder)

    } #Перенос отчетности в папки Outlook

    fMoveUnknowMessage($oInboxFolder, $oInboxItem) {

        $oTargetFolder = $oInboxFolder.Folders.Item("N2")
        [void]$oInboxItem.Move($oTargetFolder)

    } #Перенос неизвестного письма в папку unknow Outlook

    [string]fSaveAttachment($sReportsPath, $oAttachment, $sForm, $sInOut) {
    
        [string]$sFilePath = $sReportsPath + "\" + $sForm + "\" + $sInOut + "\" + $oAttachment.FileName
        $oAttachment.SaveAsFile($sFilePath)

        Return $sFilePath
    } #Сохранение вложения из письма, возвращает путь до сохраненного файла (строка)

    [string]fSaveMSG($oInboxItem, $sReportsPath) {
    
        $oInboxItem.SaveAs($sReportsPath)

        Return $sReportsPath

    } #Сохранение письма, возвращает путь до сохраненного файла (строка)

    fSendReportFiles($oOutlook, $sPath, $sForm, $sMailAddress) {

        $oMessage = $oOutlook.CreateItem(0)
        $oMessage.Recipients.Add($sMailAddress)
        $oMessage.Subject = $sForm
        $oMessage.Body = ""
        $oFiles = Get-ChildItem $sPath -File

        foreach($oFile in $oFiles) {

            $sFilePath = $oFile.FullName
            $oMessage.Attachments.Add($sFilePath)

        }

        $oMessage.Send()

        Start-Sleep -Seconds 15

    } #Отправка нескольких файлов

    fSendReportFile($oOutlook, $sReportPath, $sForm, $sMailAddress) {

        $oMessage = $oOutlook.CreateItem(0)
        $oMessage.Recipients.Add($sMailAddress)
        $oMessage.Subject = $sForm
        $oMessage.Body = ""
        $oMessage.Attachments.Add($sReportPath)
        
        $account  = $oOutlook.Session.Accounts | Where-Object { $_.DisplayName -eq $this.sBankSVKAddress }
        [Microsoft.Office.Interop.Outlook.MailItem].InvokeMember("SendUsingAccount",[System.Reflection.BindingFlags]::SetProperty,$null,$oMessage,$account)
        $oMessage.Send()

        Start-Sleep -Seconds 30

    } #Отправка файла в ЦБ через аккаунт svkN1

    fSendNBKIFile($oOutlook, $sReportPath, $sForm, $sMailAddress) {

        $oMessage = $oOutlook.CreateItem(0)
        $oMessage.Recipients.Add($sMailAddress)
        $oMessage.Subject = $sForm
        $oMessage.Body = ""
        $oFiles = Get-ChildItem $sReportPath -File
        
        foreach($oFile in $oFiles) {
            
            $sFilePath = $oFile.FullName
            $oMessage.Attachments.Add($sFilePath)
        }

        $account  = $oOutlook.Session.Accounts | Where-Object { $_.DisplayName -eq $this.sBankNBKIAddress }
        [Microsoft.Office.Interop.Outlook.MailItem].InvokeMember("SendUsingAccount",[System.Reflection.BindingFlags]::SetProperty,$null,$oMessage,$account)
        $oMessage.Send()
        
        Start-Sleep -Seconds 30

    } #Отправка файла в ЦБ через аккаунт bki_mail

    fSendReplyFile($oOutlook, $sReplyFile, $sForm, $sReportsRole) {

        [string[]]$aReportsRoles = $sReportsRole.Split("#")

        $oMessage = $oOutlook.CreateItem(0)
        
        for($i1=($aReportsRoles.count-1);$i1 -ge 0;$i1--) {
        
            $oMessage.Recipients.Add(("ROReports" + $($aReportsRoles[$i1]) + "@resocreditbank.ru"))
        
        }
        
        $oMessage.Subject = "Получен файл: " + $sForm
        $oMessage.Body = "Получен файл: " + $sForm
        $oMessage.Attachments.Add($sReplyFile)
        $oMessage.Send()
        
        Start-Sleep -Seconds 5

    } #Отправка полученного из ЦБ файла ответственной группе рассылки

    fSendReportConfirm($oOutlook, $sForm, $sReportsRole) {
    
        [string[]]$aReportsRoles = $sReportsRole.Split("#")
     
        $oMessage = $oOutlook.CreateItem(0)
     
        for($i1=($aReportsRoles.count-1);$i1 -ge 0;$i1--) {
     
            $oMessage.Recipients.Add(("ROReports" + $($aReportsRoles[$i1]) + "@resocreditbank.ru"))
        
        }
     
        $oMessage.Subject = "Подтверждение отправки: " + $sForm
        $oMessage.Body = "Файл отправлен."
        $oMessage.Send()
     
        Start-Sleep -Seconds 15
    
    } #Отправка подтверждения отправки ответственной группе рассылки

    fSendReportError($oOutlook, $sForm, $sReportsRole) {

        [string[]]$aReportsRoles = $sReportsRole.Split("#")
        
        $oMessage = $oOutlook.CreateItem(0)
        
        for($i1=($aReportsRoles.count-1);$i1 -ge 0;$i1--) {
        
            $oMessage.Recipients.Add(("ROReports" + $($aReportsRoles[$i1]) + "@resocreditbank.ru"))
        
        }
        
        $oMessage.Subject = "Ошибка отправки: " + $sForm
        $oMessage.Body = "Нет файла."
        $oMessage.Send()
        
        Start-Sleep -Seconds 15
    
    } #Ошибка отправки - нет файла ответственной группе рассылки

    fSendErorFile($oOutlook, $sForm, $sReportsRole, $res) {
    
        [string[]]$aReportsRoles = $sReportsRole.Split("#")

        $oMessage = $oOutlook.CreateItem(0)

        for($i1=($aReportsRoles.count-1);$i1 -ge 0;$i1--) {

            $oMessage.Recipients.Add(("ROReports" + $($aReportsRoles[$i1]) + "@resocreditbank.ru"))
        }

        $oMessage.Subject = "Ошибка отправки: " + $sForm
        $oMessage.Body = "Файл не отправлен. " + $res
        $oMessage.Send()

        Start-Sleep -Seconds 15

        $oMessage = $oOutlook.CreateItem(0)
        $oMessage.Recipients.Add("it@resocreditbank.ru")
        $oMessage.Subject = "Ошибка отправки: " + $sForm
        $oMessage.Body = "Файл не отправлен. " + $res
        $oMessage.Send()

        Start-Sleep -Seconds 15
    } #Отправка файла с ошибкой ответственной группе рассылки и IT

}