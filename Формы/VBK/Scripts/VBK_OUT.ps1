Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[SendErrors]$SendErr = [SendErrors]::New()

[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$FormName = Split-Path -Path $sRootFormDir -Leaf
[string]$sOutFormDir = $sRootFormDir + "\OUT"
[string]$sErrDir = $sOutFormDir + "\ERR"
[string]$sLogDir = $sOutFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

$sFilesInDir = Get-ChildItem $sOutFormDir -File

$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null

foreach($sFile in $sFilesInDir) {

    [string]$sFileName = $sFile.Name

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки файла $sFileName"  | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос на обработку файла $sFileName"  | Tee-Object $MainLogPath -Append

    [string]$sReportPath = $sFile.FullName
    [string]$sReportPath = Split-Path -Path $sReportPath -Parent
    [string]$sProcessPath = $Path_to_Folders.sPath_to_local_process + "\" + $FormName + "\OUT"
    [string]$sSendPath = $sReportPath + "\FORSEND"

    $File_operations.fMoveFile($sFile.FullName, $sProcessPath) 

    [string]$sProcessReportPath = $sProcessPath + "\" + $sFile.Name
    [string]$sExtension = $sFile.Extension.ToUpper()

    $sLog = $currentLogDir + "\" + $sFileName + ".txt"
    New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

    if(!($sExtension.Contains("XML")) -and !($sExtension.Contains("ARJ"))) {

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Подпись и шифрование файла $sFileName"  | Tee-Object $MainLogPath -Append
        $Crypt_funcs.fEncryptReportGTUsig($sProcessReportPath, $sLog)

        $content = Get-Content $sLog
        if($content -match "Ошибка"){
            
            $errName = $sFileName + ".err"
            Rename-Item $sLog $errName
            $errFile = $currentLogDir + "\" + $errName
            $SendErr.SendMail($errFile, $sProcessReportPath, "$($Name.Name)", "Ошибка обработки исходящего сообщения $($Name.Name). Файл не отправлен!")
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку ERR файла $sFileName"  | Tee-Object $MainLogPath -Append
            Move-Item $errFile $errLogDir
            Move-Item $sProcessReportPath $sErrDir

        } else {

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку отправки файла $sFileName"  | Tee-Object $MainLogPath -Append
            $File_operations.fCopyFileToFolder($sProcessReportPath, $sSendPath)

        }

    } else {

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Уже зашифрован в KlikoMsg XML-файл $sFileName"  | Tee-Object $MainLogPath -Append
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку отправки файла $sFileName"  | Tee-Object $MainLogPath -Append
        $File_operations.fCopyFileToFolder($sProcessReportPath, $sSendPath)

    }

    
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $sFileName"  | Tee-Object $MainLogPath -Append
    $File_operations.fMoveFileToARC($sProcessReportPath)

}

Remove-Item($sRootFormDir + "\" + $Name.Name)

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("$($Name.Name)", "Ошибка обработки исходящего сообщения $($Name.Name)", "$err")

    }
    $Error.Clear()

}