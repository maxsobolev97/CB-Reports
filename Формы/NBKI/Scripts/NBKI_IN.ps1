Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Run_exe]$Run_exe = [Run_exe]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[SendErrors]$SendErr = [SendErrors]::New()

[string]$sARJ = $File_operations.sARJ
[string[]]$aVFDImages = $Run_exe.aVFDImages
[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$FormName = Split-Path -Path $sRootFormDir -Leaf
[string]$sOutFormDir = $sRootFormDir + "\IN"
[string]$ProccesFormDir = $Path_to_Folders.sPath_to_local_process + "\" + $FormName + "\IN"
[string]$ForsendFormDir = $sOutFormDir + "\INBOX"
[string]$sNBKIinWorkFolder = $Path_to_Folders.sNBKIinWorkFolder
[string]$sNBKIoutWorkFolder = $Path_to_Folders.sNBKIoutWorkFolder
[string]$sNBKIfolderIN = $Path_to_Folders.sNBKIfolderIN
[string]$sErrDir = $sOutFormDir + "\ERR"
[string]$sLogDir = $sOutFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"


$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name

$FilesInDir = Get-ChildItem $sOutFormDir -File

Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Монтирую ключ NBKI" | Tee-Object $MainLogPath -Append
$Run_exe.fMountVFDImage($aVFDImages[2])

foreach($FileInDir in $FilesInDir) {
    [string]$sFileName = $FileInDir.Name
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос на обработку файла $sFileName" | Tee-Object $MainLogPath -Append
    $File_operations.fMoveFile($FileInDir.FullName, $ProccesFormDir)
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование файла для обработки $sFileName" | Tee-Object $MainLogPath -Append
    $sFilePath = $ProccesFormDir + "\" + $FileInDir.Name
    $File_operations.fCopyFileToFolder($sFilePath, $sNBKIinWorkFolder)
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос полученного файла в архив $sFileName" | Tee-Object $MainLogPath -Append
    $File_operations.fMoveFileToARC($sFilePath)
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Подготовка файла для отправки $sFileName" | Tee-Object $MainLogPath -Append
    $Crypt_funcs.fDEncryptNBKI()
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Копирование в ИБСО файла $sFileName" | Tee-Object $MainLogPath -Append
    Remove-Item C:\ScriptsNBKI\_out\*.p7s
    move C:\ScriptsNBKI\_out\* $sNBKIfolderIN
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Очистка рабочей директории $sFileName" | Tee-Object $MainLogPath -Append
    $File_operations.fDelFiles($sNBKIoutWorkFolder)
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки файла $sFileName" | Tee-Object $MainLogPath -Append
}

Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Монтирую ключ Signatura" | Tee-Object $MainLogPath -Append
$Run_exe.fMountVFDImage($aVFDImages[0])

Remove-Item($sRootFormDir + "\" + $Name.Name)

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("$($Name.Name)", "Ошибка обработки входящего сообщения $($Name.Name)", "$err")
    
    }
    $Error.Clear()

}
