Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

function isFilesInDirs{

    Param(

        $sReportsPath,
        $sPS

    )

    $sReportsDirsSVK = Get-ChildItem $sReportsPath -Directory -Exclude "CIK","NBKI","FNS"
    RunScripts $sReportsDirsSVK

    $ProcessPS = (Get-Process powershell).Count
    if ($ProcessPS -gt 4){

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Выполняется обработка отчетности ЦБ, обработка отчетности CIK и NBKI будет начата по завершеннии"

    } else {

        $sReportsDirsCIK = Get-ChildItem $sReportsPath -Directory -Filter "CIK"
        RunScripts $sReportsDirsCIK

    }

    $ProcessPS = (Get-Process powershell).Count
    if ($ProcessPS -gt 4){

        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Выполняется обработка отчетности CIK, отчетность NBKI будет начата по завершеннии"

    } else {

        $sReportsDirsNBKI = Get-ChildItem $sReportsPath -Directory -Filter "NBKI"
        RunScripts $sReportsDirsNBKI

    }

} # Опрос папок и вызов обрабатывающих скриптов

function RunScripts{

    Param($sReportsDirs)

    foreach($sReportDir in $sReportsDirs){

        $forWatch = Get-ChildItem $sReportDir.FullName | Where-Object {$_.Name -eq $sReportDir.Name}

        if($forWatch -eq $null){
            
            [string]$FolderName = $sReportDir.Name
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Просматриваю папку IN: $FolderName"
            $sReportDirIn = $sReportDir.FullName + "\IN"
            $sFilesInDir = Get-ChildItem $sReportDirIn -File

            if($sFilesInDir -ne $null){
            
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обнаружен файл в папке IN $FolderName"
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начинаю обработку поступившего файла"
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Запуск скрипта обработки $FolderName"
                $sScriptName = $sReportDir.FullName + "\Scripts\" + $sReportDir.Name + "_IN.ps1"
                if(($sReportDir.Name -eq "NBKI") -or ($sReportDir.Name -eq "CIK")){
                    Start-Process -FilePath $sPS $sScriptName -Wait
                } else {
                    Start-Process -FilePath $sPS $sScriptName
                }
            
            }

            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Просматриваю папку OUT: $FolderName"
            $sReportDirOut = $sReportDir.FullName + "\OUT"
            $sFilesInDir = Get-ChildItem $sReportDirOut -File

            if($sFilesInDir -ne $null){
            
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обнаружен файл в папке OUT $FolderName"
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начинаю обработку поступившего файла"
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Запуск скрипта обработки $FolderName"
                $sScriptName = $sReportDir.FullName + "\Scripts\" + $sReportDir.Name + "_OUT.ps1"
                if(($sReportDir.Name -eq "NBKI") -or ($sReportDir.Name -eq "CIK")){
                    Start-Process -FilePath $sPS $sScriptName -Wait
                } else {
                    Start-Process -FilePath $sPS $sScriptName
                }

            }
    
        } else {
        
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Еще не завершена обработка данных формы $($sReportDir.Name)"
        
        }

    }

} # Запуск обрабатывающих скриптов

[Run_exe]$Run_exe = [Run_exe]::New()
$sPS = $Run_exe.sPS
[string[]]$aVFDImages = $Run_exe.aVFDImages
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
$sReportsPath = $Path_to_Folders.sReportsPath
[SendErrors]$SendErr = [SendErrors]::New()

$CryptErr = "W:\example\ErrorsInCrypTool.ps1"

Start-Job -FilePath "W:\example\ErrorsInCrypTool.ps1" -Name "CryptErrors"

$sPath_mifns_serv = $Path_to_Folders.sPath_mifns_serv
$sPath_mifns_serv_out = $sPath_mifns_serv + "\out"
$sReportsPathMIFNS = $sReportsPath + "\MIFNS\OUT\"


$Run_exe.fMountVFDImage($aVFDImages[0])

while($true){

    isFilesInDirs $sReportsPath $sPS
    [System.GC]::Collect()

    if($Error){
        $err = $Error.Item(0).ToString()
        $SendErr.SendMail("Dirs_Event_Robot1", "Ошибка опроса папок Dirs_Event_Robot1!", "$err")
        $Error.Clear()
    }

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Пауза 10 секунд"
    Start-Sleep -Seconds 10

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Попытка отправки справок MIFNS"
    $files = Get-ChildItem $sPath_mifns_serv_out -File
    $date = Get-Date

    foreach($file in $files){

        $TimeToCopy = $date - $file.CreationTime
        if($TimeToCopy.Hours -ge 1){
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Перенос файла $file.name"
            Move-Item $file.FullName $sReportsPathMIFNS
        }

    }

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Пауза 10 секунд"
    Start-Sleep -Seconds 10


}