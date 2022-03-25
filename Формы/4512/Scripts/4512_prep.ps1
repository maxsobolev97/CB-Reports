Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[SendErrors]$SendErr = [SendErrors]::New()
[Run_exe]$Run_exe = [Run_exe]::New()

[string]$sARJ = $File_operations.sARJ
[string[]]$aVFDImages = $Run_exe.aVFDImages

[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$sOutFormDir = $sRootFormDir + "\OUT"
[string]$sSendPath = $sOutFormDir + "\FORSEND"
[string]$sArcPath = $sOutFormDir + "\ARC"
$Name = Get-Item $sRootFormDir
[string]$sProcessPath = $Path_to_Folders.sPath_to_local_process + "\" +$Name.Name + "\OUT"

[string]$sErrDir = $sOutFormDir + "\ERR"
[string]$sLogDir = $sOutFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

New-Item -Path $sRootFormDir -Name $Name.Name
New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null

function CopyReports($sReportDir, $sToDir){

    $sPath = "$sReportDir\$sCurrentDay"
    $bExist = Test-Path $sPath

    if($bExist){

        $oFiles = Get-ChildItem "$sReportDir\$sCurrentDay" -File

    } else{

        $oFiles = Get-ChildItem $sReportDir -File

    }

    foreach($oFile in $oFiles){
    
        Copy-Item $oFile.FullName $sToDir
    
    }

}

function MoveReports($sReportDir, $sToDir){

    $oFiles = Get-ChildItem $sReportDir -File

    foreach($oFile in $oFiles){
    
        Move-Item $oFile.FullName $sToDir
    
    }

}

function PrepareReport($sReportFolder){

    CopyReports $sReportFolder $sOutFormDir

    CopyReports $sOutFormDir $sArcPath

    MoveReports $sOutFormDir $sProcessPath

    $Run_exe.fMountVFDImage($aVFDImages[1])

    $oFilesToSig = Get-ChildItem $sProcessPath -File

    foreach($oFileToSig in $oFilesToSig){

        $sFilePath = $oFileToSig.FullName
        $Crypt_funcs.fCrypt4512u($sFilePath, $sProcessPath)

    }

    $Run_exe.fMountVFDImage($aVFDImages[0])

    foreach($oFileToSig in $oFilesToSig){

        $sLog = $currentLogDir + "\" + $oFileToSig.Name + ".txt"
        New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

        $sReportPath = $oFileToSig.FullName
        $Crypt_funcs.fCompressReportFSGzip($sReportPath, $sLog)
        $Crypt_funcs.fEncryptVKsig($sReportPath, $sLog)

    }

    $oPreparedReport = Get-ChildItem $sReportPath -File

    $sPathToPrepared = $sReportFolder + ", example\" + $sCurrentDay

    New-Item $sPathToPrepared -ItemType Directory

    MoveReports $sProcessPath $sPathToPrepared

}

$sCurrentDay = Get-Date -Format "dd.MM.yyyy"

$sDirEKVT = "R:\example"
$sDirESVT = "R:\example"
$sDirEKKR = "R:\example"
$sDirESKR = "R:\example"

PrepareReport $sDirEKVT
PrepareReport $sDirESVT
PrepareReport $sDirEKKR
PrepareReport $sDirESKR

$sFilesInDir = Get-ChildItem $sOutFormDir -File

Remove-Item($sRootFormDir + "\" + $Name.Name)

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("$($Name.Name)", "Ошибка обработки входящего сообщения $($Name.Name)", "$err")
    
    }
    $Error.Clear()

}
