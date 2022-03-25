Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Run_exe]$Run_exe = [Run_exe]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[SendErrors]$SendErr = [SendErrors]::New()

$sFNSFolderIn = $Path_to_Folders.sFNSFolderIn
[string[]]$aVFDImages = $Run_exe.aVFDImages
[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$FormName = Split-Path -Path $sRootFormDir -Leaf
[string]$sINFormDir = $sRootFormDir + "\IN"
[string]$sErrDir = $sINFormDir + "\ERR"
[string]$sLogDir = $sRootFormDir + "\IN\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

Function FNSfromCB1{

    Param(
        
        $sFiles

    )

    foreach($sFile in $sFiles){

        [string]$FileName = $sFile.Name
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование файла в папку на обработку $FileName" | Tee-Object $MainLogPath -Append
    
        [string]$sReportPath = $sFile.FullName
        [string]$sReportPath = Split-Path -Path $sReportPath -Parent
        [string]$sProcessPath = $Path_to_Folders.sPath_to_local_process + "\" + $FormName + "\IN"
        [string]$sSendPath = $sReportPath + "\INBOX"

        $File_operations.fMoveFile($sFile.FullName, $sProcessPath)

        [string]$sProcessReportPath = $sProcessPath + "\" + $sFile.Name

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

    }

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обработка полученных файлов" | Tee-Object $MainLogPath -Append
    
    $oReplyFiles = Get-ChildItem $sProcessPath

    foreach($oReplyFile in $oReplyFiles) {

        [string]$FileName = $oReplyFile.Name
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки поступившего файла $FileName" | Tee-Object $MainLogPath -Append
        $sReplyPath = $oReplyFile.FullName
        $sLog = $currentLogDir + "\" + $FileName + ".txt"
        New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null
        
        if($oReplyFile.Name -like "*.VRB"){
        
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Расшифровка файла $FileName" | Tee-Object $MainLogPath -Append
                        
            $Crypt_funcs.fDecryptReportFSsig($sReplyPath, $sLog)
        
        }

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Снятие подписи с файла $FileName" | Tee-Object $MainLogPath -Append
                   
        $Crypt_funcs.fRemoveKAReportFSsig($sReplyPath, $sLog)
        $content = Get-Content $sLog
        if($content -match "Ошибка"){
                            
            $errName = $oReplyFile.Name + ".err"
            Rename-Item $sLog $errName
            
        }
                        
    }

    $errFiles = Get-ChildItem $currentLogDir -Filter "*.err"
    foreach($errFile in $errFiles){
        
        $errorReport = $sProcessPath + "\" + $errFile.BaseName
        $SendErr.SendMail($errorReport, $errFile.FullName, "FNS", "FNS440 ошибка обработки входящих!")
        Copy-Item $errorReport $sErrDir
        Remove-Item $errorReport
        Move-Item $errFile.FullName $errLogDir
    
    }

    $ReplyFiles = Get-ChildItem $sProcessPath

    foreach($ReplyFile in $ReplyFiles) {

        [string]$FileName = $ReplyFile.Name
        $ReplyPath = $ReplyFile.FullName
    
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование файла на сервер $FileName" | Tee-Object $MainLogPath -Append
        Copy-Item $ReplyPath $sFNSFolderIn     

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $FileName" | Tee-Object $MainLogPath -Append
        Start-Sleep -Seconds 2
        $File_operations.fMoveFileToARC($ReplyPath)

    }
    
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки поступивших файлов FNS" | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date) "
    
}

$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name -ErrorAction SilentlyContinue | Out-Null

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null

$FindForsend = (Get-ChildItem $sINFormDir | Where-Object{$_.Name -eq "forsend.arj"})
if($FindForsend -ne $null){
    
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обнаружен нераспакованный файл forsend.arj, ожидание распаковки" | Tee-Object $MainLogPath -Append
    Start-Sleep 11
}

$sFilesInDir = Get-ChildItem $sINFormDir -File

FNSfromCB1 $sFilesInDir

Remove-Item($sRootFormDir + "\" + $Name.Name) -ErrorAction SilentlyContinue | Out-Null

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){
        
        $SendErr.SendMail("FNS_IN", "Ошибка обработки входящих FNS_IN!", "$err")
    }

    $Error.Clear()

}

