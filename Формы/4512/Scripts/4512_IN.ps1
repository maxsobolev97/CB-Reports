Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

[File_operations]$File_operations = [File_operations]::New()
[Crypt_funcs]$Crypt_funcs = [Crypt_funcs]::New()
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
[SendErrors]$SendErr = [SendErrors]::New()
[Run_exe]$Run_exe = [Run_exe]::New()

[string]$sRootFormDir = Split-Path -Path $PSScriptRoot -Parent
[string]$sInFormDir = $sRootFormDir + "\IN"
[string]$sSendPath = $sInFormDir + "\INBOX"
$Name = Get-Item $sRootFormDir
[string]$sProcessPath = $Path_to_Folders.sPath_to_local_process + "\" +$Name.Name + "\IN"

[string]$kwt_ekvt_path = "R:\example"
[string]$kwt_esvt_path = "R:\example"
[string]$uwd_ekvt_path = "R:\example"
[string]$uwd_esvt_path = "R:\example"
[string]$kwt_ekkr_path = "R:\example"
[string]$kwt_eskr_path = "R:\example"
[string]$uwd_ekkr_path = "R:\example"
[string]$uwd_eskr_path = "R:\example"

[string]$sErrDir = $sInFormDir + "\ERR"
[string]$sLogDir = $sInFormDir + "\LOG"
[string]$currentLogDir = $sLogDir + "\" + $(Get-Date -Format 'yyyy.MM.dd')
[string]$errLogDir = $currentLogDir + "\ErrLog"
[string]$MainLogPath = $currentLogDir + "\!MainLog.txt"

function CopyReportToEachFolder{
    
    Param($FilePath, $TargetFolder)

    $date = Get-Date -Format "dd.MM.yyyy"
    $pathToReport = $TargetFolder + "\" + $date
    New-Item $pathToReport -ItemType Directory -ErrorAction SilentlyContinue
    $File_operations.fCopyFileToFolder($FilePath, $pathToReport)

}

New-Item $currentLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $errLogDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item $MainLogPath -ItemType File -ErrorAction SilentlyContinue | Out-Null


New-Item -Path $sRootFormDir -Name $Name.Name

$FindForsend = (Get-ChildItem $sINFormDir | Where-Object{$_.Name -eq "forsend.arj"})
if($FindForsend -ne $null){
    
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обнаружен нераспакованный файл forsend.arj, ожидание распаковки" | Tee-Object $MainLogPath -Append
    Start-Sleep 11

}

$sFilesInDir = Get-ChildItem $sInFormDir -File

foreach($sFile in $sFilesInDir) {

    [string]$sFileName = $sFile.Name

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начало обработки файла $sFileName" | Tee-Object $MainLogPath -Append
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос в папку обработки файла $sFileName" | Tee-Object $MainLogPath -Append
    
    [string]$sReportPath = $sFile.FullName

    $File_operations.fMoveFile($sReportPath, $sProcessPath) 
    $filePath = $sProcessPath + "\" + $sFileName
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка дополнительной распаковки полученных файлов" | Tee-Object $MainLogPath -Append

    if($sFile.Extension.ToUpper() -eq ".ARJ"){
    
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Распаковка файла $sFileName" | Tee-Object $MainLogPath -Append
    
        $Run_exe.fUnzipReport($filePath)
        Start-Sleep -Seconds 1
        $File_operations.fDelFile($filePath)
        Start-Sleep -Seconds 1
    
        } else {
    
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Распаковка файла $sFileName не требуется." | Tee-Object $MainLogPath -Append
    
        }

}

$reports = Get-ChildItem $sProcessPath

foreach($report in $reports){
    
    $sFileName = $report.Name
    $sProcessReportPath = $report.FullName

    $sLog = $currentLogDir + "\" + $sFileName + ".txt"
    New-Item $sLog -ItemType File -ErrorAction SilentlyContinue | Out-Null

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Снятие подписи с файла $sProcessReportPath" | Tee-Object $MainLogPath -Append
    $Crypt_funcs.fRemoveKAReportFSsig($sProcessReportPath, $sLog)

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Определение типа поступившего файла $sFileName" | Tee-Object $MainLogPath -Append

    if($sFileName -match "UVDCKR" -and $report.Extension.ToUpper() -match "XML"){

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является уведомлением о приеме архивного файла example" | Tee-Object $MainLogPath -Append

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос уведомления $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
        CopyReportToEachFolder $sProcessReportPath $uwd_ekkr_path

    } elseif($sFileName -match "UVPSKR" -and $report.Extension.ToUpper() -match "XML"){

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является уведомлением о приеме архивного файла example" | Tee-Object $MainLogPath -Append
    
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос уведомления $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
        CopyReportToEachFolder $sProcessReportPath $uwd_eskr_path

    } elseif($sFileName -match "UVDCEI" -and $report.Extension.ToUpper() -match "XML"){

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является уведомлением о приеме архивного файла example" | Tee-Object $MainLogPath -Append
    
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос уведомления $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
        CopyReportToEachFolder $sProcessReportPath $uwd_ekvt_path

    } elseif($sFileName -match "UVPSEI" -and $report.Extension.ToUpper() -match "XML"){

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является уведомлением о приеме архивного файла example" | Tee-Object $MainLogPath -Append

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос уведомления $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
        CopyReportToEachFolder $sProcessReportPath $uwd_esvt_path

    } elseif($sFileName -notmatch "UVPSEI" -and $sFileName -notmatch "UVDCEI" -and $sFileName -notmatch "UVPSKR" -and $sFileName -notmatch "UVDCKR" -and $report.Extension.ToUpper() -match "XML") {

        [xml]$xmlFile = Get-Content $sProcessReportPath
    
        $sReportType = $xmlFile.KVIT.REZ_ES

        if($sReportType.ToUpper() -match "ЭКВТ"){
    
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является квитанцией о приеме файла example" | Tee-Object $MainLogPath -Append
        
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос квитанции $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
            CopyReportToEachFolder $sProcessReportPath $kwt_ekvt_path

        } elseif($sReportType.ToUpper() -match "ЭСВТ"){
    
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является квитанцией о приеме файла example" | Tee-Object $MainLogPath -Append

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос квитанции $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
            CopyReportToEachFolder $sProcessReportPath $kwt_esvt_path

        } elseif($sReportType.ToUpper() -match "ЭККР"){
    
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является квитанцией о приеме файла example" | Tee-Object $MainLogPath -Append

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос квитанции $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
            CopyReportToEachFolder $sProcessReportPath $kwt_ekkr_path

        } elseif($sReportType.ToUpper() -match "ЭСКР"){
    
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    $sFileName - Является квитанцией о приеме файла example" | Tee-Object $MainLogPath -Append

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос квитанции $sFileName в папку на диск R" | Tee-Object $MainLogPath -Append
            CopyReportToEachFolder $sProcessReportPath $kwt_eskr_path

        } 

    }

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос на отправку оператору файла $sFileName" | Tee-Object $MainLogPath -Append
    $File_operations.fCopyFileToFolder($sProcessReportPath, $sSendPath)

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Архивирование файла $sFileName" | Tee-Object $MainLogPath -Append
    $File_operations.fMoveFileToARC($sProcessReportPath)

}


Remove-Item($sRootFormDir + "\" + $Name.Name)

if($Error){

    $err = $Error.Item(0).ToString()
    if($err -notmatch "!MainLog"){

        $SendErr.SendMail("$($Name.Name)", "Ошибка обработки входящего сообщения $($Name.Name)", "$err")
    
    }
    $Error.Clear()

}
