Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[SendErrors]$SendErr = [SendErrors]::New()

[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$sInFormDir = $sRootFormDir + "\IN"
[string]$sErrDir = $sInFormDir + "\ERR"
[string]$sLogDir = $sInFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

$sFilesInDir = Get-ChildItem $sInFormDir -File

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null

$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name

foreach($sFile in $sFilesInDir) {

    [string]$sFileName = $sFile.Name

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки файла $sFileName" | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку обработки файла $sFileName" | Tee-Object $MainLogPath -Append

    [string]$sReportPath = $sFile.FullName
    [string]$sReportPath = Split-Path -Path $sReportPath -Parent
    [string]$FormName = Split-Path -Path $sReportPath -Parent
    [string]$FormName = Split-Path -Path $FormName -Leaf
    [string]$sProcessPath = $Path_to_Folders.sPath_to_local_process + "\" +$FormName + "\IN"
    [string]$sSendPath = $sReportPath + "\INBOX"

    $File_operations.fMoveFile($sFile.FullName, $sProcessPath) 
    [string]$sProcessReportPath = $sProcessPath + "\" + $sFile.Name
    [string]$sExtension = $sFile.Extension.ToUpper()

    $sLog = $currentLogDir + "\" + $sFileName + ".txt"
    New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

    if(!($sExtension.Contains("XML"))) {

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Снятие подписи с файла $sProcessReportPath" | Tee-Object $MainLogPath -Append
        $Crypt_funcs.fRemoveKAReportBRsig($sProcessReportPath, $sLog)

        $content = Get-Content $sLog
        if($content -match "Ошибка"){
            
            $errName = $sFileName + ".err"
            Rename-Item $sLog $errName
            $errFile = $currentLogDir + "\" + $errName
            $SendErr.SendMail($errFile, $sProcessReportPath, "$($Name.Name)", "Ошибка обработки входящего сообщения $($Name.Name)")
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку ERR файла $sFileName"  | Tee-Object $MainLogPath -Append
            Move-Item $errFile $errLogDir
            Move-Item $sProcessReportPath $sErrDir

        } else {
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку отправки файла $sFileName" | Tee-Object $MainLogPath -Append
            $File_operations.fCopyFileToFolder($sProcessReportPath, $sSendPath)

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $sFileName" | Tee-Object $MainLogPath -Append
            $File_operations.fMoveFileToARC($sProcessReportPath)
        
        }

    } else {

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Расшифровка и снятие КА не требуется для файла $sFileName" | Tee-Object $MainLogPath -Append
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку отправки файла $sFileName" | Tee-Object $MainLogPath -Append
        $File_operations.fCopyFileToFolder($sProcessReportPath, $sSendPath)

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $sFileName" | Tee-Object $MainLogPath -Append
        $File_operations.fMoveFileToARC($sProcessReportPath)

    }

}


Remove-Item($sRootFormDir + "\" + $Name.Name)

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("$($Name.Name)", "Ошибка обработки входящего сообщения $($Name.Name)", "$err")
    
    }
    $Error.Clear()

}
