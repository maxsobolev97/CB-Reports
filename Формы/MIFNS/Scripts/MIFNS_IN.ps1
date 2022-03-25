Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Run_exe]$Run_exe = [Run_exe]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[SendErrors]$SendErr = [SendErrors]::New()

$sPath_mifns_serv = $Path_to_Folders.sPath_mifns_serv
[string[]]$aVFDImages = $Run_exe.aVFDImages
[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$FormName = Split-Path -Path $sRootFormDir -Leaf
[string]$sINFormDir = $sRootFormDir + "\IN"
[string]$sProcessPath = $Path_to_Folders.sPath_to_local_process + "\" + $FormName + "\IN"
[string]$sSendPath = $sReportPath + "\INBOX"
[string]$sErrDir = $sINFormDir + "\ERR"
[string]$sLogDir = $sINFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

$sFilesInDir = Get-ChildItem $sINFormDir -File

Function MIFNSfromCB1{

    Param(
        
        $sFile

    )

    [string]$FileName = $sFile.Name
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки поступившего файла $FileName" | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование файла в папку на обработку" | Tee-Object $MainLogPath -Append
    
    [string]$sReportPath = $sFile.FullName
    [string]$sReportPath = Split-Path -Path $sReportPath -Parent
    
    $File_operations.fMoveFile($sFile.FullName, $sProcessPath)

    [string]$sProcessReportPath = $sProcessPath + "\" + $sFile.Name

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Первичная распаковка полученного файлов" | Tee-Object $MainLogPath -Append
    $Run_exe.fUnzipReport($sProcessReportPath)

    Start-Sleep -Seconds 1

    $File_operations.fDelFile($sProcessReportPath)

    Start-Sleep -Seconds 1

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка дополнительной распаковки полученных файлов" | Tee-Object $MainLogPath -Append

    $oReplyFiles = Get-ChildItem $sProcessPath -Filter "*.arj"

    foreach($oReplyFile in $oReplyFiles) {

        [string]$FileName = $oReplyFile.Name
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Распаковка файла $FileName" | Tee-Object $MainLogPath -Append
        $sReplyPath = $oReplyFile.FullName
        $Run_exe.fUnzipReport($sReplyPath)
        Start-Sleep -Seconds 1
        $File_operations.fDelFile($sReplyPath)
        Start-Sleep -Seconds 1

    }

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обработка полученных файлов" | Tee-Object $MainLogPath -Append
    
    $oReplyFiles = Get-ChildItem $sProcessPath -Filter "*.xml"

    foreach($oReplyFile in $oReplyFiles) {

        [string]$FileName = $oReplyFile.Name
        $sReplyPath = $oReplyFile.FullName
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Снятие подписи с файла $FileName" | Tee-Object $MainLogPath -Append

        $sLog = $currentLogDir + "\" + $FileName + ".txt"
        New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null
        $Crypt_funcs.fRemoveKAReportFSsig($sReplyPath, $sLog)

        $content = Get-Content $sLog
        if($content -match "Ошибка"){
            
            $errName = $oReplyFile.Name + ".err"
            Rename-Item $sLog $errName
            $errFile = $currentLogDir + "\" + $errName
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку ERR файла $($oReplyFile.Name)"  | Tee-Object $MainLogPath -Append
            $SendErr.SendMail($errFile, $oReplyFile.FullName, "MIFNS", "MIFNS ошибка обработки входящих!")
            Move-Item $errFile $errLogDir
            Move-Item $oReplyFile.FullName $sErrDir

        } else {

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование файла на сервер $FileName" | Tee-Object $MainLogPath -Append
            $sPath_mifns_serv_out = $sPath_mifns_serv + "\in"
            copy $sReplyPath $sPath_mifns_serv_out
        
            if($oReplyFile.Name -like "UV*"){
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование файла на отправку оператору $FileName" | Tee-Object $MainLogPath -Append
                $sFormInboxDir = $sINFormDir + "\INBOX"
                copy $sReplyPath $sFormInboxDir
            }

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $FileName" | Tee-Object $MainLogPath -Append
            Start-Sleep -Seconds 5
            $File_operations.fMoveFileToARC($sReplyPath)
        
        }

    }
    
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступивших файлов MIFNS" | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date) "
    
}

$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null

foreach($sFileInDir in $sFilesInDir){

    MIFNSfromCB1 $sFileInDir

}

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("MIFNS_IN", "Ошибка обработки входящих MIFNS_IN!", "$err")

    }
    $Error.Clear()

}

Remove-Item($sRootFormDir + "\" + $Name.Name)