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
[string]$sOutFormDir = $sRootFormDir + "\OUT"
[string]$ProccesFormDir = $Path_to_Folders.sPath_to_local_process + "\" + $FormName + "\OUT"
[string]$sCIKFolderIn = $Path_to_Folders.sCIKFolderIn
[string]$sErrDir = $sOutFormDir + "\ERR"
[string]$sLogDir = $sOutFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"


$Name = Get-Item $sRootFormDir
New-Item -Path $sRootFormDir -Name $Name.Name

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null

$sFilesInDirectory = Get-ChildItem $sOutFormDir -File

Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Монтирую ключ CIK" | Tee-Object $MainLogPath -Append
$Run_exe.fMountVFDImage($aVFDImages[1])

foreach($sFileInDirectory in $sFilesInDirectory){

    [string]$FileName = $sFileInDirectory.Name
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обработка файла" | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос файла на обработку $FileName" | Tee-Object $MainLogPath -Append
    $File_operations.fMoveFile($sFileInDirectory.FullName, $ProccesFormDir)

}

$sFilesProcessDirectory = Get-ChildItem $ProccesFormDir -File

foreach($sFileProcessDirectory in $sFilesProcessDirectory){

    [string]$FileName = $sFileProcessDirectory.Name
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Расшифровка файла $FileName" | Tee-Object $MainLogPath -Append
    $Crypt_funcs.fDecryptCIK($sFileProcessDirectory.FullName)

}

$sFilesProcessDirectory = Get-ChildItem $ProccesFormDir -File

foreach($sFileProcessDirectory in $sFilesProcessDirectory){

    $FileName = $sFileProcessDirectory.Name
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос готового файла в example $FileName" | Tee-Object $MainLogPath -Append
    $File_operations.fCopyFileToFolder($sFileProcessDirectory.FullName, $sCIKFolderIn)
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование готового файла $FileName" | Tee-Object $MainLogPath -Append
    $File_operations.fMoveFileToARC($sFileProcessDirectory.FullName)

}

Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Монтирую ключ Signatura" | Tee-Object $MainLogPath -Append
$Run_exe.fMountVFDImage($aVFDImages[0])

Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Завершение обработки" | Tee-Object $MainLogPath -Append

Remove-Item($sRootFormDir + "\" + $Name.Name)

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("$($Name.Name)", "Ошибка обработки исходящего сообщения $($Name.Name)", "$err")

    }
    $Error.Clear()

}