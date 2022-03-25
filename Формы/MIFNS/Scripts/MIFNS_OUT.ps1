Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\Library\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Run_exe]$Run_exe = [Run_exe]::New()
[SendErrors]$SendErr = [SendErrors]::New()

[string[]]$aVFDImages = $Run_exe.aVFDImages
[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$FormName = Split-Path -Path $sRootFormDir -Leaf
[string]$sOutFormDir = $sRootFormDir + "\OUT"
[string]$sForSend = $sRootFormDir + "\OUT\FORSEND"
[string]$ProccesFormDir = $Path_to_Folders.sPath_to_local_process + "\" + $FormName + "\OUT"
[string]$sPath_mifns_serv = $Path_to_Folders.sPath_mifns_serv
[string]$sPath_mifns_serv_files = $sPath_mifns_serv + "\out"
[string]$sErrDir = $sOutFormDir + "\ERR"
[string]$sLogDir = $sOutFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

Function fEncryptMIFNS {
    Param(
        
        $sFilesInDir,
        $sProcessPath,
        $sOutFormDir

    )

    if($null -ne $sFilesInDir){

        foreach($sFile in $sFilesInDir){

            [string]$sName = $sFile.Name
            $sTimeToChange = $sFile.LastWriteTime
            $sDate = Get-Date
            $sRezDate = $sDate - $sTimeToChange

            [string]$sReportArchPath = ""
            
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обработка файлов" | Tee-Object $MainLogPath -Append
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование файла на обработку $sName"

            [string]$sReportPath = $sFile.FullName
            [string]$sReportPath = Split-Path -Path $sReportPath -Parent
            [string]$sSendPath = $sReportPath + "\FORSEND"

            $File_operations.fMoveFile($sFile.FullName, $sProcessPath) 

            [string]$sProcessReportPath = $sProcessPath + "\" + $sFile.Name

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Печать файла $sName" | Tee-Object $MainLogPath -Append
            [string]$sNewName = $sProcessPath + "\" + $sFile.BaseName + ".txt"
            Rename-Item -Path $sProcessReportPath -NewName $sNewName
            Start-Process –FilePath $sNewName –Verb Print -Wait
            Rename-Item -Path $sNewName -NewName $sProcessReportPath

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Подпись файла $sName"

            $sLog = $currentLogDir + "\" + $sName + ".txt"

            New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null
            $Crypt_funcs.fSetKAReportFSsig($sProcessReportPath, $sLog)

            Start-Sleep 1
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сжатие файла отчёта Gzip $sName" | Tee-Object $MainLogPath -Append
            $sLog = $currentLogDir + "\" + $sName + ".txt"
            New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

            $Crypt_funcs.fCompressReportFSGzip($sProcessReportPath, $sLog)
            Start-Sleep 1
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Шифрование файла отчёта $sName" | Tee-Object $MainLogPath -Append
            $sLog = $currentLogDir + "\" + $sName + ".txt"
            New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

            $Crypt_funcs.fEncryptMIFNSsig($sProcessReportPath, $sLog)
            Start-Sleep 1
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Добавление в архивный файл файла $sName" | Tee-Object $MainLogPath -Append

            $content = Get-Content $sLog
            if($content -match "Ошибка"){
            
                $errName = $oReplyFile.Name + ".err"
                Rename-Item $sLog $errName
            
            }
                
        }

    }

}#Шифрование и подпись файлов отчетов

Function fArchReportMIFNS1 {
    Param(
        $sReportPath,
        $sOutFormDir,
        $sProcessPath
    )

    [Run_exe]$Run_exe = [Run_exe]::New()
    [File_operations]$File_operations = [File_operations]::New()
    $sARJ = $File_operations.sARJ

    [int]$iNumReports = fGetNumTodayReports ($sOutFormDir + "\ARC")
    [string]$sPathProcess = $sProcessPath
    [string]$sDate = Get-Date -format yyMMdd
    [string]$sNumReport = ($iNumReports + 1).ToString()
    [string]$sArchReportPath = $ProccesFormDir + "\AN25451" + $sDate + "000" + $sNumReport + ".arj"
    if($iNumReports -gt 8) {
        $sArchReportPath = $ProccesFormDir + "\AN25451" + $sDate + "00" + $sNumReport + ".arj"
    }
    if(Test-Path ($ProccesFormDir + "\*_700.xml")) {
        $sArchReportPath = $ProccesFormDir + "\BN25451" + $sDate + "000" + $sNumReport + ".arj"
        if($iNumReports -gt 8) {
            $sArchReportPath = $ProccesFormDir + "\BN25451" + $sDate + "00" + $sNumReport + ".arj"
        }
    }
    $Run_exe.fRunExe($sARJ, " a -e " + $sArchReportPath + " " + $sReportPath.FullName) | Out-Null

    Return $sArchReportPath
} #Сборка архивного файла тип А/В

Function fEncryptArchMIFNS {
    Param(
    
        $sProcessPath,
        $sForSend,
        $sOutFormDir,
        $ProccesFormDir

    )

    if($null -ne $sProcessPath){

        foreach($sFile in $sProcessPath){

            [string]$sReportPath = $sFile.FullName
            [string]$sReportPath = Split-Path -Path $sReportPath -Parent
            

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Подпись архивного файла с отчётами" | Tee-Object $MainLogPath -Append
            
            $sFileName = $sFile.Name
            $sLog = $currentLogDir + "\" + $sFileName + ".txt"
            New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

            $Crypt_funcs.fSetKAReportFSsig($sFile.FullName, $sLog)
            Start-Sleep 1
            $content = Get-Content $sLog
            if($content -match "Ошибка"){
            
                $errName = $oReplyFile.Name + ".err"
                Rename-Item $sLog $errName
            
            }
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Упаковка архивного файла с отчётами в архивный файл" | Tee-Object $MainLogPath -Append
            

            

        }

    }

} #Подпись архивного файла

Function fArchReportMIFNS2 {
    Param(
        $sReportPath,
        $sOutFormDir,
        $ProccesFormDir
    )

    [Run_exe]$Run_exe = [Run_exe]::New()
    [File_operations]$File_operations = [File_operations]::New()
    $sARJ = $File_operations.sARJ

    [int]$iNumReports = fGetNumTodayReports ($sOutFormDir + "\ARC")
    [string]$sPathProcess = $ProccesFormDir
    [string]$sDate = Get-Date -format yyMMdd
    [string]$sNumReport = ($iNumReports + 1).ToString()
    [string]$sArchReportPath = $ProccesFormDir + "\MIFNS2.00" + $sNumReport
    if($iNumReports -gt 8) {
        $sArchReportPath = $ProccesFormDir + "\MIFNS2.0" + $sNumReport
    }
    $Run_exe.fRunExe($sARJ, " a -e " + $sArchReportPath + " " + $sPathProcess + "\*.arj") | Out-Null

    $sArchReportPathToreturn = Get-Item $sArchReportPath
    $sPathToReturn = $ProccesFormDir + "\" + $sArchReportPathToreturn.Name
    $File_operations.fMoveFileToARC($sReportPath.FullName)
    Return $sPathToReturn
} #Сборка транспортного файла

Function fGetNumTodayReports {
    Param(
        $sPath
    )

    [string]$sDate = Get-Date -format yyyyMMdd
    [int]$iNum = (Get-ChildItem $sPath -filter ("MIFNS2_" + $sDate + "*")).Count
    $iNum = $iNum + 1
    Return $iNum

} #Получить номер прошлого архивного файла

$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null
    
$sFilesInDir700 = Get-ChildItem $sOutFormDir -Filter *_700.xml
      
if($sFilesInDir700 -ne $null){

    Start-Sleep 3
    $sFilesInDir700s = Get-ChildItem $sOutFormDir -Filter *_700.xml
    fEncryptMIFNS $sFilesInDir700s $ProccesFormDir $sOutFormDir
    $errFiles = Get-ChildItem $currentLogDir -Filter "*.err"
    foreach($errFile in $errFiles){
        
        $errorReport = $ProccesFormDir + "\" + $errFile.BaseName
        $SendErr.SendMail($errorReport, $errFile.FullName, "MIFNS", "MIFNS ошибка обработки исходящих!")
        Copy-Item $errorReport $sErrDir
        Remove-Item $errorReport
        Move-Item $errFile.FullName $errLogDir
    
    }
    Start-Sleep 3
    $sProcessReportsPath = Get-ChildItem $ProccesFormDir -Filter *_700.xml
    foreach($sProcessReportPath in $sProcessReportsPath){
        $sReportArchPath = fArchReportMIFNS1 $sProcessReportPath $sOutFormDir $sProcessPath
        $File_operations.fMoveFileToARC($sProcessReportPath.FullName)
    }
    Start-Sleep 3
    $sArchInProcessDir700s = Get-ChildItem $ProccesFormDir -File
    fEncryptArchMIFNS $sArchInProcessDir700s $sForSend, $sOutFormDir $ProccesFormDir
    $errFiles = Get-ChildItem $currentLogDir -Filter "*.err"
    foreach($errFile in $errFiles){
        
        $errorReport = $ProccesFormDir + "\" + $errFile.BaseName
        $SendErr.SendMail($errorReport, $errFile.FullName, "MIFNS", "MIFNS ошибка обработки исходящих!")
        Copy-Item $errorReport $sErrDir
        Remove-Item $errorReport
        Move-Item $errFile.FullName $errLogDir
    
    }
    Start-Sleep 3

    $sArchsInProcessDir700s = Get-ChildItem $ProccesFormDir -File
    foreach($sArchInProcessDir700s in $sArchsInProcessDir700s){
        $sReportArchPath2 = fArchReportMIFNS2 $sArchInProcessDir700s $sOutFormDir $ProccesFormDir
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос готового архивного файла в папку отправки" | Tee-Object $MainLogPath -Append
        $File_operations.fCopyFileToFolder($sReportArchPath2, $sForSend)
        $File_operations.fMoveFileToARC($sReportArchPath2)

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки" | Tee-Object $MainLogPath -Append
    }
    Start-Sleep 5
}
    
$sFilesInDirwthout700 = Get-ChildItem $sOutFormDir -Exclude ARC,FORSEND,PROCESS,ERR,LOG,*700.xml

if($sFilesInDirwthout700 -ne $null){

    Start-Sleep 3
    $sFilesInDirwthout700s = Get-ChildItem $sOutFormDir -Exclude ARC,FORSEND,PROCESS,ERR,LOG,*700.xml
    fEncryptMIFNS $sFilesInDirwthout700s $ProccesFormDir $sOutFormDir
    $errFiles = Get-ChildItem $currentLogDir -Filter "*.err"
    foreach($errFile in $errFiles){
        
        $errorReport = $ProccesFormDir + "\" + $errFile.BaseName
        $SendErr.SendMail($errorReport, $errFile.FullName, "MIFNS", "MIFNS ошибка обработки исходящих! Файл не отправлен!")
        Copy-Item $errorReport $sErrDir
        Remove-Item $errorReport
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку ERR файла $($errFile.FullName)"  | Tee-Object $MainLogPath -Append
        Move-Item $errFile.FullName $errLogDir
    
    }
    Start-Sleep 3
    $sProcessReportsPath = Get-ChildItem $ProccesFormDir -Exclude ARC,FORSEND,PROCESS,ERR,LOG,*700.xml
    foreach($sProcessReportPath in $sProcessReportsPath){
        $sReportArchPath = fArchReportMIFNS1 $sProcessReportPath $sOutFormDir $sProcessPath
        $File_operations.fMoveFileToARC($sProcessReportPath.FullName)
    }
    Start-Sleep 3
    $sArchInProcessDirwthout700s = Get-ChildItem $ProccesFormDir -File
    fEncryptArchMIFNS $sArchInProcessDirwthout700s $sForSend, $sOutFormDir $ProccesFormDir
    $errFiles = Get-ChildItem $currentLogDir -Filter "*.err"
    foreach($errFile in $errFiles){
        
        $errorReport = $ProccesFormDir + "\" + $errFile.BaseName
        $SendErr.SendMail($errorReport, $errFile.FullName, "MIFNS", "MIFNS ошибка обработки исходящих! Файл не отправлен!")
        Copy-Item $errorReport $sErrDir
        Remove-Item $errorReport
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку ERR файла $($errFile.FullName)"  | Tee-Object $MainLogPath -Append
        Move-Item $errFile.FullName $errLogDir
    
    }
    Start-Sleep 3

    $sArchsInProcessDirwthout700s = Get-ChildItem $ProccesFormDir -File
    foreach($sArchInProcessDirwthout700s in $sArchsInProcessDirwthout700s){
        $sReportArchPath2 = fArchReportMIFNS2 $sArchInProcessDirwthout700s $sOutFormDir $ProccesFormDir
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос готового архивного файла в папку отправки" | Tee-Object $MainLogPath -Append
        $File_operations.fCopyFileToFolder($sReportArchPath2, $sForSend)
        $File_operations.fMoveFileToARC($sReportArchPath2)

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки" | Tee-Object $MainLogPath -Append
    }
    Start-Sleep 5
}

Remove-Item($sRootFormDir + "\" + $Name.Name)

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("MIFNS_OUT", "Ошибка обработки исходящих MIFNS_OUT!", "$err")

    }
    $Error.Clear()

}