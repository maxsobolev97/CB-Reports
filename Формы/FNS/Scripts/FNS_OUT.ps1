Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Run_exe]$Run_exe = [Run_exe]::New()
[SendErrors]$SendErr = [SendErrors]::New()

[string]$sARJ = $File_operations.sARJ
[string[]]$aVFDImages = $Run_exe.aVFDImages
[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$FormName = Split-Path -Path $sRootFormDir -Leaf
[string]$sOutFormDir = $sRootFormDir + "\OUT"
[string]$ProccesFormDir = $Path_to_Folders.sPath_to_local_process + "\" + $FormName + "\OUT"
[string]$sFNSFolderOut = $Path_to_Folders.sFNSFolderOut
[string]$sErrDir = $sOutFormDir + "\ERR"
[string]$sLogDir = $sOutFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

Function SendFNS() {

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки ответов по 440-П" | Tee-Object $MainLogPath -Append
    [string]$sFNSPath = $ProccesFormDir
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос файлов ответа для обработки" | Tee-Object $MainLogPath -Append
    $sFNSFilesInOUTdir = $sFNSFolderOut + "\*"
    $File_operations.fMoveFile($sFNSFilesInOUTdir,$sOutFormDir)
    $localFiles = Get-ChildItem $sOutFormDir -File

    foreach($localFile in $localFiles){
    
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос файла $($localFile.Name) в локальную папку" | Tee-Object $MainLogPath -Append
        Move-Item $localFile.FullName $ProccesFormDir

    }

    $oFNSFiles = Get-ChildItem $sFNSPath -File

    if($oFNSFiles.Count -gt 0) {

        foreach($oFNSFile in $oFNSFiles) {
            
            [string]$FileName = $oFNSFile.Name
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Подпись файла $FileName" | Tee-Object $MainLogPath -Append
            $sFilePath = $oFNSFile.FullName
            $sLog = $currentLogDir + "\" + $FileName + ".txt"
            New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

            $Crypt_funcs.fSetKAReportFSsig($sFilePath, $sLog)

            if((($oFNSFile.Extension.ToUpper() -eq ".XML") -or ($oFNSFile.Extension.ToUpper() -eq ".VRB")) -and (-not($sFilePath.Contains("PB1_"))) -and (-not($sFilePath.Contains("PB2_")))) {
                
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Сжатие файла Gzip $FileName" | Tee-Object $MainLogPath -Append
                $Crypt_funcs.fCompressReportFSGzip($sFilePath, $sLog)
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Шифрование файла $FileName" | Tee-Object $MainLogPath -Append
                $Crypt_funcs.fEncryptFNSsig($sFilePath, $sLog)

            }

            $content = Get-Content $sLog
            if($sLog.Contains("PB1_") -or $sLog.Contains("PB2_")){
            
                if(($content -match "КА установлен") -and !($content -match "Ошибка")){
            
                    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Проверка результата обработки файла $FileName прошла успешно!" | Tee-Object $MainLogPath -Append

                } else {
                    
                    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Проверка результата обработки файла $FileName завершилась неудачей!" | Tee-Object $MainLogPath -Append       
                    $errName = $oFNSFile.Name + ".err"
                    Rename-Item $sLog $errName
            
                }
            
            } else {

                if(($content -match "КА установлен") -and ($content -match "Файл сжат") -and ($content -match "Зашифрован на абонента") -and !($content -match "Ошибка")){
                    
                    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Проверка результата обработки файла $FileName прошла успешно!" | Tee-Object $MainLogPath -Append
            
                } else {

                    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Проверка результата обработки файла $FileName завершилась неудачей!" | Tee-Object $MainLogPath -Append        
                    $errName = $oFNSFile.Name + ".err"
                    Rename-Item $sLog $errName
            
                }

            }

        }

        $errFiles = Get-ChildItem $currentLogDir -Filter "*.err"
        foreach($errFile in $errFiles){
        
            $errorReport = $ProccesFormDir + "\" + $errFile.BaseName
            $SendErr.SendMail($errorReport, $errFile.FullName, "FNS", "FNS440 ошибка обработки исходящих! Файл не отправлен!")
            Copy-Item $errorReport $sErrDir
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку ERR файла $($errFile.FullName)"  | Tee-Object $MainLogPath -Append
            Remove-Item $errorReport
            Move-Item $errFile.FullName $errLogDir
    
        }


        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Упаковка файлов ответа в архивные файлы" | Tee-Object $MainLogPath -Append
        fArchFNS1 $ProccesFormDir
 
        $oFNSArc1Files = Get-ChildItem $ProccesFormDir -Filter "*.arj"

        foreach($oFNSArc1File in $oFNSArc1Files) {

            [string]$FileName = $oFNSArc1File.Name
            $sLog = $currentLogDir + "\" + $FileName + ".txt"
            New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Подпись файла $FileName" | Tee-Object $MainLogPath -Append
            $sFileArc1Path = $oFNSArc1File.FullName
            $Crypt_funcs.fSetKAReportFSsig($sFileArc1Path, $sLog)

        }
        
        $errFiles = Get-ChildItem $currentLogDir -Filter "*.err"
        foreach($errFile in $errFiles){
        
            $errorReport = $ProccesFormDir + "\" + $errFile.BaseName
            $SendErr.SendMail($errorReport, $errFile.FullName, "FNS", "FNS440 ошибка обработки исходящих!")
            Copy-Item $errorReport $sErrDir
            Remove-Item $errorReport
            Move-Item $errFile.FullName $errLogDir
    
        }


        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Упаковка архивных файлов в транспортные файлы" | Tee-Object $MainLogPath -Append
        fArchFNS2 $ProccesFormDir
        $oFNSArc2Files = Get-ChildItem $ProccesFormDir -File
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос готовых архивных файлов на отправку" | Tee-Object $MainLogPath -Append
        $sForSendDir = $sOutFormDir + "\FORSEND"
        $File_operations.fCopyFilesToFolder($ProccesFormDir, $sForSendDir)
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Очистка директории обработки" | Tee-Object $MainLogPath -Append

        foreach($oFNSArc2File in $oFNSArc2Files) {
            
            [string]$FileName = $oFNSArc2File.Name
            $sFileArc2Path = $oFNSArc2File.FullName
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $FileName" | Tee-Object $MainLogPath -Append
            $File_operations.fMoveFileToARC($sFileArc2Path)
            
        }
        
    }

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки ответов по 440-П" | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date) " | Tee-Object $MainLogPath -Append
    
}

Function ErrorProcessing(){


    $filesInError = Get-ChildItem $sErrDir
    $filesInLogDir = Get-ChildItem $errLogDir

    if($filesInError.Count -gt $filesInLogDir.Count){
    
        $SendErr.SendMail("FNS_OUT", "Ошибка повторной обработки исходящих FNS_OUT!", "Отсутствуют необходимые файлы логов для повторной обработки, требуется ручная обработка!")
    
    } elseif($filesInError.Count -lt $filesInLogDir.Count){
    
        $SendErr.SendMail("FNS_OUT", "Ошибка повторной обработки исходящих FNS_OUT!", "Количество ошибочных логов превышает количество файлов в папке ошибок, требуется ручная обработка!")
    
    } elseif($filesInError.Count -eq $filesInError.Count -and $filesInError.Count -ne 0){
        
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки ошибок ответов по 440-П" | Tee-Object $MainLogPath -Append

        $filesInError = Get-ChildItem $sErrDir -Exclude "*.arj"
        foreach($fileInError in $filesInError){
            
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос ошибочного файла $fileInError в локальную папку для повторной обработки" | Tee-Object $MainLogPath -Append
            Move-Item $fileInError.FullName $ProccesFormDir

            $localName = $ProccesFormDir + "\" + $fileInError.Name
            $fileInErrorLog = $errLogDir + "\" + $fileInError.Name + ".err"
            $sLog = $currentLogDir + "\" + $fileInError.Name + ".txt"

            $logContent = Get-Content $fileInErrorLog
            $logKA = if($logContent -match "КА установлен"){'1'}else{'0'}
            $logZIP = if($logContent -match "Файл сжат"){'1'}else{'0'}
            $logEncrypt = if($logContent -match "Зашифрован на абонента"){'1'}else{'0'}

            $sCondition = $logKA + $logZIP + $logEncrypt

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Статус файла $fileInError - $sCondition" | Tee-Object $MainLogPath -Append

            switch($sCondition){
            
                '100'{
                      
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Снятие КА с файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fRemoveKAReportFSsig($localName, $sLog)
                      
                }
                '110'{

                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Распаковка файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fDecompressReportFSGzip($localName, $sLog)
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Снятие КА с файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fRemoveKAReportFSsig($localName, $sLog)
                }
                '010'{

                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Распаковка файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fDecompressReportFSGzip($localName, $sLog)

                }
                '011'{
                
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Распаковка файла $fileInError" | Tee-Object $MainLogPath -Append
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Расшифровка файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fDecryptReportFSsig($localName, $sLog)
                        
                }
                '101'{
                
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Расшифровка файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fDecryptReportFSsig($localName, $sLog)
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Снятие КА с файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fRemoveKAReportFSsig($localName, $sLog)
                        
                }
                '001'{
                
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Расшифровка файла $fileInError" | Tee-Object $MainLogPath -Append
                        $Crypt_funcs.fDecryptReportFSsig($localName, $sLog)
                        
                }
                '111'{
                        
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Исключительная ситуация для файла $fileInError, файл прошел все проверки, требуется ручная обработка!" | Tee-Object $MainLogPath -Append
                        $SendErr.SendMail("FNS_OUT", "Ошибка повторной обработки исходящих FNS_OUT!", "Исключительная ситуация для файла $fileInError, файл прошел все проверки, требуется ручная обработка!")

                }
                '000'{
                
                        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Действий над файлом $fileInError не требуется" | Tee-Object $MainLogPath -Append
                                        
                }
            
            }

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Удаление лог-файла $fileInErrorLog с ошибками" | Tee-Object $MainLogPath -Append
            Remove-Item $fileInErrorLog
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос чистого файла $fileInError в сетевую папку для повторной обработки" | Tee-Object $MainLogPath -Append
            Move-Item $localName $sOutFormDir

        }


        $SendErr.SendMail("FNS_OUT", "Успех повторной обработки исходящих FNS_OUT", "Файлы успешно обработаны!")
    
    }


}

Function fArchFNS1 {
    Param(
        $sPath
    )

    [int]$iNumReports = fGetNumTodayReportsFNS ($sOutFormDir + "\ARC")
    [string]$sDate = Get-Date -format yyyyMMdd
    [string]$sNumReport = ""
    [string]$sArchPath = ""
    [string]$sZeros = ""
    [string]$sFilePath = ""
    [int]$iCount = 0
    $oFiles = Get-ChildItem $sPath\* -Include @("*.xml","*.vrb")
    foreach($oFile in $oFiles) {

        $iCount++
        $sNumReport = ($iNumReports + 1).ToString()
        if($iNumReports -lt 9) {
            $sZeros = "0000"
        } elseif($iNumReports -lt 99) {
            $sZeros = "000"
        } else {
            $sZeros = "00"
        }
        $sArchPath = $ProccesFormDir + "\AFN_4525451_MIFNS00_" + $sDate + "_" + $sZeros + $sNumReport + ".arj"
        $sFilePath = $oFile.FullName
        $Run_Exe.fRunExe($sARJ, " a -e " + $sArchPath + " " + $sFilePath) | Out-Null
        $File_operations.fMoveFileToARC($sFilePath)

        
        if(((Get-Item $sArchPath).length -gt 5MB) -or ($iCount -eq 49)) {
            $iNumReports = $iNumReports + 1
            $iCount = 0
        }
    }
    Start-Sleep 3
}

Function fArchFNS2 {
    Param(
        $sPath
    )

    [int]$iNumReports = fGetNumTodayReportsFNS ($sOutFormDir + "\ARC")
    [string]$sDate = Get-Date -format yyyyMMdd
    [string]$sNumReport = ""
    [string]$sArchPath = ""
    [string]$sZeros = ""
    [string]$sFilePath = ""
    $oFiles = Get-ChildItem $sPath -Filter "*.arj"
    foreach($oFile in $oFiles) {
        
        $sNumReport = ($iNumReports + 1).ToString()
        if($iNumReports -lt 9) {
            $sZeros = "00"
        } elseif($iNumReports -lt 99) {
            $sZeros = "0"
        } else {
            $sZeros = ""
        }
        $sArchPath = $ProccesFormDir + "\FNS440." + $sZeros + $sNumReport
        $sFilePath = $oFile.FullName
        $Run_exe.fRunExe($sARJ, " a -e " + $sArchPath + " " + $sFilePath) | Out-Null
        
        $sFilePathToARC = $ProccesFormDir + "\" + $oFile.Name
        $File_operations.fMoveFileToARC($sFilePathToARC)
        $iNumReports = $iNumReports + 1
    }
    Start-Sleep 3

}

Function fGetNumTodayReportsFNS {
    Param(
        $sPath
    )

    [string]$sDate = Get-Date -format yyyyMMdd
    [int]$iNum = (Get-ChildItem $sPath -filter ("FNS440_" + $sDate + "*")).Count
    $iNum = $iNum + 1
    Return $iNum
}

$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name -ErrorAction SilentlyContinue | Out-Null

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null

SendFNS
Start-Sleep 3
ErrorProcessing

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("FNS_OUT", "Ошибка обработки исходящих FNS_OUT!", "$err")

    }
    $Error.Clear()

}

Remove-Item($sRootFormDir + "\" + $Name.Name) -ErrorAction SilentlyContinue | Out-Null